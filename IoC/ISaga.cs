using System;
using System.Collections;
using System.Collections.Generic;

namespace IoC
{
    public interface ISaga
    {
        Guid Id { get; }
        int Version { get; }

        void Transition(object message);

        IEnumerable<IVersionedEvent> GetUncommittedEvents();
        void ClearUncommittedEvents();

        IEnumerable<IVersionedEvent> GetUndispatchedMessages();
        void ClearUndispatchedMessages();
    }
}