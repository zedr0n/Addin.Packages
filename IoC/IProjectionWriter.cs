using System;
using NodaTime;

namespace IoC
{
    public interface IProjectionWriter<in TId,TProjection>
        where TProjection : class
    {
        void Add(TId id, TProjection projection);
        /// <summary>
        /// Add the projection with the specified timestamp
        /// </summary>
        /// <param name="id">Aggregate id</param>
        /// <param name="projection">Projection instance</param>
        /// <param name="asOf">Timestamp of projection to add</param>
        void Add(TId id, TProjection projection, Instant asOf);
        /// <summary>
        /// Update the projections after asOf time
        /// </summary>
        /// <param name="id">Aggregate id</param>
        /// <param name="action">Update action</param>
        /// <param name="asOf">Earliest timestamp of projections to update</param>
        /// <exception cref="InvalidOperationException">No projections satisfying the time criteria</exception>
        void UpdateOrThrow(TId id, Action<TProjection> action, Instant asOf);
    }

    public interface IProjectionWriter<TProjection> : IProjectionWriter<Guid, TProjection>
        where TProjection : class
    {
        
    }
}