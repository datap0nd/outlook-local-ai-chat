using System;
using System.Threading;

namespace OutlookLocalAIChat.Security
{
    public sealed class OneShotDraftAuthorization
    {
        private int _available;
        private int _consumed;
        private int _created;

        public OneShotDraftAuthorization(bool authorized)
        {
            WasAuthorized = authorized;
            _available = authorized ? 1 : 0;
        }

        public bool WasAuthorized { get; }

        public bool IsConsumed
        {
            get
            {
                return Volatile.Read(ref _consumed) == 1;
            }
        }

        public bool IsCreated
        {
            get
            {
                return Volatile.Read(ref _created) == 1;
            }
        }

        public bool TryConsume()
        {
            if (Interlocked.CompareExchange(
                ref _available,
                0,
                1) != 1)
            {
                return false;
            }

            Volatile.Write(ref _consumed, 1);
            return true;
        }

        internal void MarkCreated()
        {
            if (!IsConsumed)
            {
                throw new InvalidOperationException(
                    "Draft permission must be consumed before success is recorded.");
            }

            Volatile.Write(ref _created, 1);
        }
    }
}
