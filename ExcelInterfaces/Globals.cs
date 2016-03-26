﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection.Emit;
using System.Xml.Serialization;

namespace ExcelInterfaces
{
    public interface IPublicObject
    {
        string Handle { get; set; }
        string Type { get; set; }
    }

    public class Error : ApplicationException
    {
        public Error(string message) :
            base(message)
        {
        }
    }
    public class PropertyMissing : Error
    {
        public PropertyMissing(string paramName) :
            base("Parameter missing : " + paramName)
        {
        }
    }
    public class ObjectMissing : Error
    {
        public ObjectMissing(string handle) :
            base("Object missing : " + handle)
        {
        }
    }

    public static class ObjectExtensions
    {
        public static Public<T> ToPublic<T>(this T obj, string handle) where T : class
        {
            return new Public<T>(handle, obj);
        }
    }

    public class Public<T> : IPublicObject, IEquatable<Public<T>> where T : class
    {
        public T Instance;
        public Dictionary<Type,IPublicObject> Children = new Dictionary<Type, IPublicObject>(); 

        public static Public<T> This(string handle)
        {
            IPublicObject publicObject;
            if (!Globals.TryGetItem(handle, out publicObject))
                throw new ObjectMissing(handle);

            return publicObject as Public<T>;
        }

        public static bool TryThis(string handle, out Public<T> obj)
        {
            obj = null;
            try
            {
                obj = This(handle);
                return true;
            }
            catch (Exception)
            {
                // ignored
            }
            return false;
        } 

        public string AddChild<TChild>(TChild obj) where TChild : class
        {
            var tHandle = Globals.StripHandle(Handle);
            var child = obj.ToPublic(tHandle);

            return AddChild(child);
        }

        public string AddChild<TChild>(Public<TChild> obj) where TChild : class
        {
            IPublicObject child;
            if (!Children.TryGetValue(typeof(TChild), out child))
                Children.Add(typeof(TChild), obj);
            child = child ?? obj;

            return child.Handle;
        }

        public Public(string handle, T obj)
        {
            Instance = obj;
            Type = obj.GetType().Name;
            Handle = Globals.AddItem(Globals.StripHandle(handle), this);
        }

        public string Handle { get; set; }
        public string Type { get; set; }

        public string Serialise()
        {
            var x = new XmlSerializer(Instance.GetType());
            var sw = new StringWriter();
            var ns = new XmlSerializerNamespaces();
            ns.Add("", "");

            x.Serialize(sw, Instance, ns);
            return sw.ToString();
        }

        public static Public<T> Deserialise(string handle, string xml)
        {
            var sr = new StringReader(xml);
            var x = new XmlSerializer(typeof(T));
            var obj = (T)x.Deserialize(sr);

            return new Public<T>(handle, obj);
        }

        public bool Equals(Public<T> other)
        {
            var equatable = Instance as IEquatable<T>;
            if(equatable != null)
                return equatable.Equals(other.Instance);

            return Instance.Equals(other.Instance);
        }
    }

    public static class Globals
    {
        private static readonly Dictionary<string, IPublicObject> Items = new Dictionary<string, IPublicObject>();

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
            var tHandle = TimestampHandle(handle + "::" + obj.Type) + "::";
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
        public static bool TryGetItem<TValue>(string handle,out TValue obj) where TValue : IPublicObject
        {
            return Items.TryGetTypedValue(handle, out obj);
        }
        private static string TimestampHandle(string handle)
        {
            return handle + "::" + DateTime.Now.ToString("hh:mm:ss");
        }

        public static string StripHandle(string handle)
        {
            return handle.Split(new char[] {':'}).FirstOrDefault();
        }
    }
}
