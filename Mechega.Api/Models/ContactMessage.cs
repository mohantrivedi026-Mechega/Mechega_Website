using System;

namespace Mechega.Api.Models
{
    public class ContactMessage
    {
        public int Id { get; set; }
        public string? Name { get; set; }
        public string? Email { get; set; }
        public string? Phone { get; set; }
        public string? Subject { get; set; }
        public string? Message { get; set; }
        public DateTime ReceivedAt { get; set; } = DateTime.UtcNow;
        public bool Sent { get; set; }
        public DateTime? SentAt { get; set; }
    }
}
