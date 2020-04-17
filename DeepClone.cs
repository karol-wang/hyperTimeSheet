using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using AutoMapper;
using System.Reflection;
using Newtonsoft.Json;

namespace hyperTimeSheet
{
    public static class DeepCloneExtensions
    {
        /// <summary>
        /// 複製物件
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static T DeepClone<T>(this T obj)
        {
            using (var ms = new MemoryStream())
            {
                var formatter = new BinaryFormatter();
                formatter.Serialize(ms, obj);
                ms.Position = 0;

                return (T)formatter.Deserialize(ms);
            }
        }

        public static T DeepCloneByAutoMapping<T>(this T obj, T newObj) 
        {
            //T newObj = (T)Activator.CreateInstance(typeof(T));
            var config = new MapperConfiguration(cfg =>
            {
                cfg.CreateMap<T, T>();
            });
            var mapper = config.CreateMapper();
            mapper.Map(obj,newObj);
            
            return newObj;
        }

        public static T DeepCloneBySerializeObject<T>(this T obj)
        {
            return JsonConvert.DeserializeObject<T>(JsonConvert.SerializeObject(obj));
        }

        public static T DeepCloneByReflection<T>(this T objSource, T objTarget)
        {
            //Get the type of source object and create a new instance of that type
            Type typeSource = objSource.GetType();
            //T objTarget = (T)Activator.CreateInstance(typeof(T));
            //Get all the properties of source object 
            PropertyInfo[] propertyInfo = typeSource.GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);

            //Assign all source property to taget object 's properties
            foreach (PropertyInfo property in propertyInfo)
            {
                //Check whether property can be written to
                if (property.CanWrite)
                {
                    //check whether property type is value type, enum or string type
                    if (property.PropertyType.IsValueType || property.PropertyType.IsEnum || property.PropertyType.Equals(typeof(System.String)))
                    {
                        property.SetValue(objTarget, property.GetValue(objSource, null), null);
                    }
                    //else property type is object/complex types, so need to recursively call this method until the end of the tree is reached
                    else
                    {
                        object objPropertyValue = property.GetValue(objSource, null);
                        if (objPropertyValue == null)
                        {
                            property.SetValue(objTarget, null, null);
                        }
                        else
                        {
                            property.SetValue(objTarget, objPropertyValue.DeepCloneByReflection(objTarget), null);
                        }
                    }
                }
            }

            return objTarget;
        }
    }
}
