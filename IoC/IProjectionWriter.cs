using System;

namespace IoC
{
    public interface IProjectionWriter<in TId,TProjection>
        where TProjection : class
    {
        void Add(TId id, TProjection projection);
        void UpdateOrThrow(TId id, Action<TProjection> action);
    }

    public interface IProjectionWriter<TProjection> : IProjectionWriter<Guid, TProjection>
        where TProjection : class
    {
        
    }
}