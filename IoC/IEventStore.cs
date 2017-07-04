using System.Collections.Generic;
using NodaTime;

namespace IoC
{
    public interface IEventStore
    {
        IEnumerable<IVersionedEvent> GetEvents();
        /// <summary>
        /// Get events before upTo instant
        /// </summary>
        /// <param name="upTo"></param>
        /// <returns></returns>
        IEnumerable<IVersionedEvent> GetEvents(Instant upTo);

    }
}