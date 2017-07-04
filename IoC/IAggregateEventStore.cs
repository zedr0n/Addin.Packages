using System;
using System.Collections;

namespace IoC
{
    public interface IAggregateEventStore<T> where T : IAggregateRoot
    {
        /// <summary>
        /// Saves the aggregate events to event store and publish those to event bus
        /// </summary>
        /// <param name="aggregate">The aggregate instance</param>
        void Save(T aggregate);
        /// <summary>
        /// Rebuild the aggregate from event history extracted from Event Store
        /// </summary>
        /// <param name="id">The aggregate guid</param>
        /// <returns>Aggregate or null if no events found</returns>
        T Find(Guid id);
        /// <summary>
        /// Rebuilds the aggregate from event history extracted from Event Store
        /// </summary>
        /// <exception cref="ArgumentException">if no events found for aggregate with this id</exception>
        /// <param name="id">The aggregate guid</param>
        /// <returns></returns>
        T Get(Guid id);

    }

    public interface ISagaEventStore<T> where T : ISaga
    {
        void Save(T saga);
        T Get(Guid id);
    }
}