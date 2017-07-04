using System;
using System.Collections.Generic;

namespace IoC
{
    public interface IAggregateRoot
    {
        Guid Id { get; }
        int Version { get; }
        IEnumerable<IVersionedEvent> Events { get; }

        void ClearChanges();
    }

}