using System;
using NodaTime;

namespace IoC
{
    public interface IVersionedEvent : IEvent
    {
        int Version { get; }
        Instant Timestamp { get; }
    }

    public abstract class VersionedEvent : IVersionedEvent
    {
        public Guid SourceId { get; set; }
        public int Version { get; set; }
        public Instant Timestamp { get; set; }
    }
}