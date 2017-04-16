using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Xml.Serialization;

namespace ExcelInterfaces
{
    public class Public<T> : Public, IEquatable<Public<T>> where T : class
    {
        public T Instance { get; }
        public Dictionary<Type,IPublicObject> Children = new Dictionary<Type, IPublicObject>();
        public override object Object => Instance;

        public new static Public<T> This(string handle)
        {
            IPublicObject publicObject;
            if (!Globals.TryGetItem(handle, out publicObject))
                throw new ObjectMissing(handle);

            var instance = publicObject.GetType().GetRuntimeProperty(nameof(Instance)).GetValue(publicObject);

            //var obj =  new Public<T>(handle, instance as T);
            //Globals.SetItem(handle,obj);
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

        public Public() { }

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

    public class Public : IPublicObject
    {
        public string Handle { get; set; }
        public string Type { get; set; }
        public virtual object Object { get; } = null;

        public static IPublicObject This(string handle)
        {
            IPublicObject publicObject;
            if (!Globals.TryGetItem(handle, out publicObject))
                throw new ObjectMissing(handle);

            return publicObject;
        }

        public static string Serialise(string handle)
        {
            var oObj = This(handle);

            var x = new XmlSerializer((Type) oObj.Object.GetType());
            var sw = new StringWriter();
            var ns = new XmlSerializerNamespaces();
            ns.Add("", "");

            x.Serialize(sw, oObj.Object, ns);
            return sw.ToString();
        }

        public string Serialise()
        {
            var x = new XmlSerializer(Object.GetType());
            var sw = new StringWriter();
            var ns = new XmlSerializerNamespaces();
            ns.Add("", "");

            x.Serialize(sw, Object, ns);
            return sw.ToString();
        }
    }
}