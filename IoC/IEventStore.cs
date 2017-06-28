using System.Collections.Generic;

namespace IoC
{
    public interface IEventStore
    {
        IEnumerable<IVersionedEvent> GetEvents();
    }
}