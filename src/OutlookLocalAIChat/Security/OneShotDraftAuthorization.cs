using System;
using System.Threading;

namespace OutlookLocalAIChat.Security
{
    public sealed class OneShotDraftAuthorization
    {
        private int _available;
        private int _consumed;
        private int _created;
        private int _updated;

        public OneShotDraftAuthorization(bool authorized)
            : this(authorized, false)
        {
        }

        public OneShotDraftAuthorization(
            bool allowCreate,
            bool allowUpdate)
        {
            CanCreate = allowCreate;
            CanUpdate = allowUpdate;
            WasAuthorized = allowCreate || allowUpdate;
            _available = WasAuthorized ? 1 : 0;
        }

        public bool WasAuthorized { get; }

        public bool CanCreate { get; }

        public bool CanUpdate { get; }

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

        public bool IsUpdated
        {
            get
            {
                return Volatile.Read(ref _updated) == 1;
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

        internal void MarkUpdated()
        {
            if (!IsConsumed)
            {
                throw new InvalidOperationException(
                    "Draft permission must be consumed before success is recorded.");
            }

            Volatile.Write(ref _updated, 1);
        }
    }
}
