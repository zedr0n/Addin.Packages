using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelInterfaces
{
    public static class ObjectExtensions
    {
        public static Public<T> ToPublic<T>(this T obj, string handle) where T : class
        {
            return new Public<T>(handle, obj);
        }
    }

    public static class Globals
    {
        [ThreadStatic]
        private static Dictionary<string, IPublicObject> _items;

        private static Dictionary<string, IPublicObject> Items => _items ?? (_items = new Dictionary<string, IPublicObject>());

        public static void Reset()
        {
            Items.Clear();
        }

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
        public static string AddItem(string handle, IPublicObject obj)
        {
            var tHandle = TimestampHandle(handle + "::" + obj.Type);
            if (!Items.ContainsKey(tHandle))
                Items.Add(tHandle,obj);

            obj.Handle = tHandle;

            return tHandle;
        }
        public static IPublicObject GetItem(string handle)
        {
            IPublicObject obj;
            return TryGetItem(handle, out obj) ? obj : null;
        }
        public static void SetItem(string handle, IPublicObject obj)
        {
            if (Items.ContainsKey(handle))
                Items[handle] = obj;
            else
                throw new ArgumentException();
        }

        public static bool TryGetItem<TValue>(string handle,out TValue obj) where TValue : IPublicObject
        {
            return Items.TryGetTypedValue(handle, out obj);
        }
        private static string TimestampHandle(string handle)
        {
            return handle + "::" + DateTime.Now.ToString("hh:mm:ss.ff");
        }

        public static string StripHandle(string handle)
        {
            return handle.Split(':').FirstOrDefault();
        }
    }
}
