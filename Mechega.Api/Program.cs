using MailKit.Security;
using MimeKit;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using Microsoft.EntityFrameworkCore;
using Mechega.Api.Data;
using Mechega.Api.Models;

var builder = WebApplication.CreateBuilder(args);

// Configuration keys (set these in appsettings or environment variables in production)
// SendGrid: SendGrid:ApiKey, SendGrid:FromEmail, SendGrid:FromName, SendGrid:ToEmail

builder.Services.AddOpenApi();
builder.Services.AddEndpointsApiExplorer();

// CORS - allow Blazor dev client origin
builder.Services.AddCors(options =>
{
    options.AddPolicy("LocalDev", policy =>
    {
        policy.WithOrigins("http://localhost:5001").AllowAnyHeader().AllowAnyMethod();
    });
});

// Database
builder.Services.AddDbContext<ContactDbContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("DefaultConnection")));

// Email sender
// Register an HttpClient for SendGrid and choose implementation at runtime.
builder.Services.AddHttpClient("sendgrid");
builder.Services.AddHttpClient("graph");
builder.Services.AddSingleton<IEmailSender>(sp =>
{
    var cfg = sp.GetRequiredService<IConfiguration>();
    // Prefer Microsoft Graph if client credentials provided
    var graphClientId = cfg["Graph:ClientId"];
    var graphTenant = cfg["Graph:TenantId"];
    var graphSecret = cfg["Graph:ClientSecret"];
    if (!string.IsNullOrEmpty(graphClientId) && !string.IsNullOrEmpty(graphTenant) && !string.IsNullOrEmpty(graphSecret))
    {
        return new GraphEmailSender(cfg, sp.GetRequiredService<ILogger<GraphEmailSender>>(), sp.GetRequiredService<IHttpClientFactory>());
    }

    var apiKey = cfg["SendGrid:ApiKey"];
    if (!string.IsNullOrEmpty(apiKey))
    {
        return new SendGridEmailSender(cfg, sp.GetRequiredService<ILogger<SendGridEmailSender>>(), sp.GetRequiredService<IHttpClientFactory>());
    }

    return new SmtpEmailSender(cfg, sp.GetRequiredService<ILogger<SmtpEmailSender>>());
});

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.MapOpenApi();
}

app.UseCors("LocalDev");

app.MapPost("/api/contact", async (ContactRequest req, ContactDbContext db, IEmailSender sender, IConfiguration config) =>
{
    if (string.IsNullOrWhiteSpace(req.Name) || string.IsNullOrWhiteSpace(req.Email) || string.IsNullOrWhiteSpace(req.Message))
        return Results.BadRequest("Name, Email and Message are required.");

    var phone = string.IsNullOrWhiteSpace(req.Phone) ? null : req.Phone.Trim();
    var message = $"Name: {req.Name}\nEmail: {req.Email}\nPhone: {phone ?? "Not provided"}\n\nMessage:\n{req.Message}";
    var subject = req.Subject ?? "Website contact form";
    var companyEmail = config["Smtp:ToEmail"] ?? "info@mechega.com";
    var fromEmail = config["Smtp:FromEmail"] ?? "info@mechega.com";

    // store message in DB
    var entity = new ContactMessage
    {
        Name = req.Name,
        Email = req.Email,
        Phone = phone,
        Subject = req.Subject,
        Message = req.Message,
        ReceivedAt = DateTime.UtcNow,
        Sent = false
    };
    db.ContactMessages.Add(entity);
    await db.SaveChangesAsync();

    var internalEmailSent = await sender.SendEmailAsync(subject, message, req.Email, companyEmail, fromEmail);

    var customerMessage = $"Hi {req.Name},\n\nThank you for contacting Mechega. We have received your message and will get back to you shortly.\n\nYour message:\n{req.Message}";
    var customerAckSent = await sender.SendEmailAsync("Thank you for contacting Mechega", customerMessage, companyEmail, req.Email, fromEmail);

    var sent = internalEmailSent || customerAckSent;

    // update message record (record whether email was sent)
    entity.Sent = sent;
    if (sent)
        entity.SentAt = DateTime.UtcNow;
    await db.SaveChangesAsync();

    // Return success to the client even if email failed so messages are persisted in dev.
    return Results.Ok(new { success = true, emailSent = sent });
});

// Ensure database is created at startup (development)
using (var scope = app.Services.CreateScope())
{
    var db = scope.ServiceProvider.GetRequiredService<ContactDbContext>();
    db.Database.EnsureCreated();

    var connection = db.Database.GetDbConnection();
    if (connection.State != System.Data.ConnectionState.Open)
        connection.Open();

    using var cmd = connection.CreateCommand();
    cmd.CommandText = "SELECT COUNT(*) FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'ContactMessages' AND COLUMN_NAME = 'Phone'";
    var columnCount = Convert.ToInt32(cmd.ExecuteScalar());

    if (columnCount == 0)
    {
        using var alter = connection.CreateCommand();
        alter.CommandText = "ALTER TABLE [ContactMessages] ADD [Phone] NVARCHAR(MAX) NULL";
        alter.ExecuteNonQuery();
    }
}

app.MapGet("/api/admin/contacts", async (int? skip, int? take, ContactDbContext db) =>
{
    var q = db.ContactMessages.OrderByDescending(c => c.ReceivedAt).AsQueryable();
    if (skip.HasValue) q = q.Skip(skip.Value);
    if (take.HasValue) q = q.Take(take.Value);
    var list = await q.Select(c => new { c.Id, c.Name, c.Email, c.Phone, c.Subject, c.Message, c.ReceivedAt, c.Sent, c.SentAt }).ToListAsync();
    return Results.Ok(list);
});

// Diagnostic: show configured connection string and row count
app.MapGet("/api/admin/dbinfo", async (ContactDbContext db, IConfiguration cfg) =>
{
    var conn = cfg.GetConnectionString("DefaultConnection");
    var count = await db.ContactMessages.CountAsync();
    return Results.Ok(new { connection = conn, count });
});

app.Run();

public record ContactRequest(string? Name, string? Email, string? Message, string? Subject, string? Phone);

public interface IEmailSender
{
    Task<bool> SendEmailAsync(string subject, string content, string? replyToEmail, string? toEmail = null, string? fromEmail = null);
}

public class SmtpEmailSender : IEmailSender
{
    private readonly IConfiguration _config;
    private readonly ILogger<SmtpEmailSender> _logger;

    public SmtpEmailSender(IConfiguration config, ILogger<SmtpEmailSender> logger)
    {
        _config = config;
        _logger = logger;
    }

    public async Task<bool> SendEmailAsync(string subject, string content, string? replyToEmail, string? toEmail = null, string? fromEmail = null)
    {
        var host = _config["Smtp:Host"];
        var port = int.TryParse(_config["Smtp:Port"], out var p) ? p : 587;
        var user = _config["Smtp:Username"];
        var pass = _config["Smtp:Password"];
        fromEmail ??= _config["Smtp:FromEmail"] ?? _config["SendGrid:FromEmail"];
        var fromName = _config["Smtp:FromName"] ?? "Mechega Website";
        toEmail ??= _config["Smtp:ToEmail"] ?? _config["SendGrid:ToEmail"];
        var useSsl = bool.TryParse(_config["Smtp:UseSsl"], out var u) ? u : true;

        if (string.IsNullOrEmpty(host) || string.IsNullOrEmpty(toEmail) || string.IsNullOrEmpty(fromEmail))
        {
            _logger.LogError("SMTP configuration missing (Host/ToEmail/FromEmail)");
            return false;
        }

        try
        {
            var message = new MimeMessage();
            message.From.Add(new MailboxAddress(fromName, fromEmail));
            message.To.Add(MailboxAddress.Parse(toEmail));
            if (!string.IsNullOrEmpty(replyToEmail))
                message.ReplyTo.Add(MailboxAddress.Parse(replyToEmail));
            message.Subject = subject;
            message.Body = new TextPart("plain") { Text = content };

            using var client = new MailKit.Net.Smtp.SmtpClient();
            var socketOption = useSsl ? SecureSocketOptions.StartTls : SecureSocketOptions.Auto;
            await client.ConnectAsync(host, port, socketOption);

            if (!string.IsNullOrEmpty(user) && !string.IsNullOrEmpty(pass))
            {
                await client.AuthenticateAsync(user, pass);
            }

            await client.SendAsync(message);
            await client.DisconnectAsync(true);
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error sending email via SMTP");
            return false;
        }
    }
}

public class SendGridEmailSender : IEmailSender
{
    private readonly IConfiguration _config;
    private readonly ILogger<SendGridEmailSender> _logger;
    private readonly IHttpClientFactory _httpFactory;

    public SendGridEmailSender(IConfiguration config, ILogger<SendGridEmailSender> logger, IHttpClientFactory httpFactory)
    {
        _config = config;
        _logger = logger;
        _httpFactory = httpFactory;
    }

    public async Task<bool> SendEmailAsync(string subject, string content, string? replyToEmail, string? toEmail = null, string? fromEmail = null)
    {
        var apiKey = _config["SendGrid:ApiKey"];
        fromEmail ??= _config["SendGrid:FromEmail"] ?? _config["Smtp:FromEmail"];
        var fromName = _config["SendGrid:FromName"] ?? _config["Smtp:FromName"] ?? "Mechega Website";
        toEmail ??= _config["SendGrid:ToEmail"] ?? _config["Smtp:ToEmail"];

        if (string.IsNullOrEmpty(apiKey) || string.IsNullOrEmpty(fromEmail) || string.IsNullOrEmpty(toEmail))
        {
            _logger.LogError("SendGrid configuration missing (ApiKey/FromEmail/ToEmail)");
            return false;
        }

        try
        {
            var payload = new
            {
                personalizations = new[] {
                    new {
                        to = new[] { new { email = toEmail } },
                        subject = subject
                    }
                },
                from = new { email = fromEmail, name = fromName },
                content = new[] { new { type = "text/plain", value = content } }
            };

            if (!string.IsNullOrEmpty(replyToEmail))
            {
                var dict = JsonSerializer.Deserialize<Dictionary<string, object>>(JsonSerializer.Serialize(payload)) ?? new Dictionary<string, object>();
                dict["reply_to"] = new { email = replyToEmail };
                var json = JsonSerializer.Serialize(dict);
                using var req = new HttpRequestMessage(HttpMethod.Post, "https://api.sendgrid.com/v3/mail/send");
                req.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey);
                req.Content = new StringContent(json, Encoding.UTF8, "application/json");
                var client = _httpFactory.CreateClient("sendgrid");
                var resp = await client.SendAsync(req);
                if (resp.IsSuccessStatusCode) return true;
                var body = await resp.Content.ReadAsStringAsync();
                _logger.LogError("SendGrid send failed: {Status} {Body}", resp.StatusCode, body);
                return false;
            }

            var jsonPayload = JsonSerializer.Serialize(payload);
            using var request = new HttpRequestMessage(HttpMethod.Post, "https://api.sendgrid.com/v3/mail/send");
            request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey);
            request.Content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");

            var http = _httpFactory.CreateClient("sendgrid");
            var response = await http.SendAsync(request);
            if (response.IsSuccessStatusCode)
                return true;

            var respBody = await response.Content.ReadAsStringAsync();
            _logger.LogError("SendGrid send failed: {Status} {Body}", response.StatusCode, respBody);
            return false;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error sending email via SendGrid");
            return false;
        }
    }
}

public class GraphEmailSender : IEmailSender
{
    private readonly IConfiguration _config;
    private readonly ILogger<GraphEmailSender> _logger;
    private readonly IHttpClientFactory _httpFactory;

    public GraphEmailSender(IConfiguration config, ILogger<GraphEmailSender> logger, IHttpClientFactory httpFactory)
    {
        _config = config;
        _logger = logger;
        _httpFactory = httpFactory;
    }

    private async Task<string?> GetAccessTokenAsync()
    {
        try
        {
            var tenant = _config["Graph:TenantId"] ?? string.Empty;
            var clientId = _config["Graph:ClientId"] ?? string.Empty;
            var clientSecret = _config["Graph:ClientSecret"] ?? string.Empty;
            if (string.IsNullOrWhiteSpace(tenant) || string.IsNullOrWhiteSpace(clientId) || string.IsNullOrWhiteSpace(clientSecret))
            {
                _logger.LogError("Graph client credentials are missing (TenantId/ClientId/ClientSecret)");
                return null;
            }

            var url = $"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token";
            var body = new List<KeyValuePair<string, string>>
            {
                new KeyValuePair<string, string>("grant_type", "client_credentials"),
                new KeyValuePair<string, string>("client_id", clientId),
                new KeyValuePair<string, string>("client_secret", clientSecret),
                new KeyValuePair<string, string>("scope", "https://graph.microsoft.com/.default")
            };

            var client = _httpFactory.CreateClient("graph");
            var resp = await client.PostAsync(url, new FormUrlEncodedContent(body));
            var json = await resp.Content.ReadAsStringAsync();
            if (!resp.IsSuccessStatusCode)
            {
                _logger.LogError("Token request failed: {Status} {Body}", resp.StatusCode, json);
                return null;
            }

            using var doc = JsonDocument.Parse(json);
            if (doc.RootElement.TryGetProperty("access_token", out var tok))
                return tok.GetString();
            return null;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error acquiring Graph token");
            return null;
        }
    }

    public async Task<bool> SendEmailAsync(string subject, string content, string? replyToEmail, string? toEmail = null, string? fromEmail = null)
    {
        fromEmail ??= _config["Graph:FromEmail"] ?? _config["Smtp:FromEmail"];
        toEmail ??= _config["Graph:ToEmail"] ?? _config["Smtp:ToEmail"];

        if (string.IsNullOrEmpty(fromEmail) || string.IsNullOrEmpty(toEmail))
        {
            _logger.LogError("Graph email config missing FromEmail/ToEmail");
            return false;
        }

        var token = await GetAccessTokenAsync();
        if (string.IsNullOrEmpty(token)) return false;

        try
        {
            var msg = new
            {
                message = new
                {
                    subject = subject,
                    body = new { contentType = "Text", content = content },
                    toRecipients = new[] { new { emailAddress = new { address = toEmail } } }
                }
            };

            if (!string.IsNullOrEmpty(replyToEmail))
            {
                var dict = JsonSerializer.Deserialize<Dictionary<string, object>>(JsonSerializer.Serialize(msg)) ?? new Dictionary<string, object>();
                if (dict.TryGetValue("message", out var rawMessageObj) && rawMessageObj is not null)
                {
                    var message = JsonSerializer.Deserialize<Dictionary<string, object>>(rawMessageObj.ToString() ?? "{}") ?? new Dictionary<string, object>();
                    message["replyTo"] = new[] { new { emailAddress = new { address = replyToEmail } } };
                    dict["message"] = message;
                    var finalJson = JsonSerializer.Serialize(dict);
                    using var req = new HttpRequestMessage(HttpMethod.Post, $"https://graph.microsoft.com/v1.0/users/{fromEmail}/sendMail");
                    req.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
                    req.Content = new StringContent(finalJson, Encoding.UTF8, "application/json");
                    var client = _httpFactory.CreateClient("graph");
                    var resp = await client.SendAsync(req);
                    if (resp.IsSuccessStatusCode) return true;
                    var body = await resp.Content.ReadAsStringAsync();
                    _logger.LogError("Graph send failed: {Status} {Body}", resp.StatusCode, body);
                    return false;
                }
            }

            var json = JsonSerializer.Serialize(msg);
            using var request = new HttpRequestMessage(HttpMethod.Post, $"https://graph.microsoft.com/v1.0/users/{fromEmail}/sendMail");
            request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
            request.Content = new StringContent(json, Encoding.UTF8, "application/json");

            var http = _httpFactory.CreateClient("graph");
            var response = await http.SendAsync(request);
            if (response.IsSuccessStatusCode) return true;

            var respBody = await response.Content.ReadAsStringAsync();
            _logger.LogError("Graph send failed: {Status} {Body}", response.StatusCode, respBody);
            return false;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error sending email via Graph");
            return false;
        }
    }
}
