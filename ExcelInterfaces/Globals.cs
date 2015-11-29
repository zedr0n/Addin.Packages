using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml.Serialization;

namespace ExcelInterfaces
{
    public interface IPublicObject : ICloneable
    {
        string Handle { get; set; }
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

    public static class PublicObject
    {
        public static T This<T>(string handle) where T : IPublicObject
        {
            T publicObject;
            if (!Globals.TryGetItem(handle, out publicObject))
                throw new ObjectMissing(handle);

            return publicObject;
        }

        public static string WriteToXml2(this IPublicObject obj)
        {
            var fieldList = new List<FieldInfo>(obj.GetType().GetFields()
                .Where(info => info.IsPublic))
                .Where(info => info.GetValue(obj).GetType().GetInterfaces().Contains(typeof (IPublicObject)))
                .Where(info => info.FieldType != info.GetValue(obj).GetType());
            var typesList = fieldList.Select(field => field.GetValue(obj).GetType()).Distinct().ToList();

            var x = new XmlSerializer(obj.GetType(), typesList.ToArray());
            var sw = new StringWriter();
            var ns = new XmlSerializerNamespaces();
            ns.Add("", "");

            x.Serialize(sw, obj, ns);
            return sw.ToString();
        }

        public static string WriteToXml<T>(this T obj) where T : IPublicObject
        {
            var privateObject = obj.ToPrivate();

            var x = new XmlSerializer(privateObject.GetType());
            var sw = new StringWriter();
            var ns = new XmlSerializerNamespaces();
            ns.Add("", "");

            x.Serialize(sw, privateObject, ns);
            return sw.ToString();
        }

        public static T ReadFromXml<T>(string xml)
        {
            var sr = new StringReader(xml);
            var x = new XmlSerializer(typeof (T));
            return (T) x.Deserialize(sr);
        }

        private static object ToPrivate<T>(this T obj) where T : IPublicObject
        {
            return obj.Clone();
        }

        public static string ToPublic<T>(this T obj, string handle)
        {
            var o = obj as IPublicObject;
            if (o != null)
                return o.Handle;

            // get all possible derived objects
            // should be just the corresponding public type
            var allTypes = AppDomain.CurrentDomain.GetAssemblies()
                .Where(x => x.GetCustomAttributes(typeof (PublicAttribute), false).Length > 0)
                .SelectMany(x => x.GetTypes())
                .Where(x => x.GetInterfaces().Contains(typeof (IPublicObject)));
            var publicType = allTypes
                .Where(t => typeof(T).IsAssignableFrom(t))
                .ToArray().Single();

            // find the (string, baseType) constructor
            var argumentTypes = new List<Type> {typeof (string), typeof (T)};
            var constructorInfo = publicType.GetConstructor(argumentTypes.ToArray());

            // and invoke with handle
            var arguments = new List<object> {handle, obj};
            var publicObject = (IPublicObject) constructorInfo?.Invoke(arguments.ToArray());

            // enumerate all fields
            var fieldList = new List<FieldInfo>(typeof (T).GetFields()
                .Where(info => info.IsPublic));
            foreach (var fieldInfo in fieldList)
            {
                var privateFieldType = fieldInfo.FieldType;
                // get all public objects derived from the private object field type
                var publicFieldType = allTypes
                    .Where(t => t.GetInterfaces().Contains(typeof (IPublicObject)))
                    .Where(t => privateFieldType.IsAssignableFrom(t))
                    .ToArray().SingleOrDefault();

                if (publicFieldType == default(Type))
                    continue;

                // find the (string, baseType) constructor
                argumentTypes = new List<Type> {typeof (string), privateFieldType};
                constructorInfo = publicFieldType.GetConstructor(argumentTypes.ToArray());
                // and invoke with handle_${baseType.Name}
                arguments = new List<object> {handle + "_" + privateFieldType.Name, fieldInfo.GetValue(publicObject)};
                // store the constructed public object
                fieldInfo.SetValue(publicObject, constructorInfo?.Invoke(arguments.ToArray()));
            }

            return publicObject?.Handle;
        }
    }

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

        public static string AddItem(string handle, IPublicObject obj)
        {
            var tHandle = TimestampHandle(handle) + "::" + obj.GetType().BaseType?.Name + "::";
            if (!Items.ContainsKey(tHandle))
                Items.Add(tHandle, obj);

            // store timestamped handle
            obj.Handle = tHandle;

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
            return handle + "::" + DateTime.Now.ToString("mm:ss.ffff");
        }
    }
}
