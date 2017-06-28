﻿using System;
using System.Collections;
using System.Collections.Generic;

namespace IoC
{
    public interface IEventStore
    {
        IEnumerable<IVersionedEvent> GetEvents();
    }

    public interface IAggregateEventStore<T> where T : class
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
}