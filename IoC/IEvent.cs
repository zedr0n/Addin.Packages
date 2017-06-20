using System;

namespace IoC
{
    public interface IEvent
    {
        Guid SourceId { get; }
    }
}