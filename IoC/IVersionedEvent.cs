using System;

namespace IoC
{
    public interface IVersionedEvent : IEvent
    {
        int Version { get; }
    }

    public abstract class VersionedEvent : IVersionedEvent
    {
        public Guid SourceId { get; set; }
        public int Version { get; set; }
    }
}