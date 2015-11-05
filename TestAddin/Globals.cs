using System;
using System.Collections.Generic;

namespace TestAddin
{
    public static class Globals
    {
        private static readonly Dictionary<string, object> Items = new Dictionary<string, object>();

        private static bool TryGetTypedValue<TKey, TValue, TActual>(
            this IDictionary<TKey, TValue> data,
            TKey key,
            out TActual value) where TActual : TValue
        {
            TValue tmp;
            if (data.TryGetValue(key, out tmp))
            {
                value = (TActual) tmp;
                return true;
            }
            value = default(TActual);
            return false;
        }

        public static string AddItem(string handle, object obj)
        {
            var tHandle = TimestampHandle(handle) + "::" + obj.GetType().Name + "::";
            if (!Items.ContainsKey(tHandle))
                Items.Add(tHandle,obj);

            return tHandle;
        }

        public static object GetItem(string handle)
        {
            object obj;
            return TryGetItem(handle, out obj) ? obj : null;
        }

        public static bool TryGetItem<TValue>(string handle,out TValue obj)
        {
            return Items.TryGetTypedValue(handle, out obj);
        }


        private static string TimestampHandle(string handle)
        {
            return handle + "::" + DateTime.Now.ToString("HH:mm:ss");
        }
    }
}
