namespace IoC
{
    public interface IConsumer
    {
    }

    public interface IConsumer<in TEvent> : IConsumer where TEvent : IEvent
    {
        void When(TEvent e);
    }
}