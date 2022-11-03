using System;
using System.Collections.Generic;
using System.Reflection;
using System.Linq;

namespace H.Npoi.ExcelHelper
{

    /// <summary>
    /// 
    /// </summary>
    public static class ImportDataExtension
    {
        public delegate void PropertySetter<T>(T value);

        public static List<T> ToModel<T>(this List<ImportDataModel> data) where T : ImportBaseModel, new()
        {
            return data.Select(s => s.ToModel<T>()).ToList();
        }

        public static T ToModel<T>(this ImportDataModel data) where T : ImportBaseModel, new()
        {
            var result = new T();
            try
            {
                result = data.Row.ToModel<T>();
                result.RowNo = data.RowNo;
                result.ErrorMsg = data.ErrorMsg;
            }
            catch (Exception ex)
            {
                var msg = $"error at row {data.RowNo},{ex.Message}";
                result.RowNo = data.RowNo;
                result.ErrorMsg = data.ErrorMsg;
                result.ErrorMsg += string.IsNullOrEmpty(result.ErrorMsg) ? msg : ";" + msg;
            }

            return result;
        }

        public static T ToModel<T>(this List<ImportColumnModel> rowData) where T : class, new()
        {
            var result = new T();
            Type type = result.GetType();
            var members = type.GetProperties();

            foreach (var item in members)
            {
                // 获取每个成员拥有的特性
                var attribute = (ColumnPropertyAttribute)item.GetCustomAttributes().FirstOrDefault(s => s.GetType() == typeof(ColumnPropertyAttribute));
                if (attribute != null)
                {
                    var value = rowData.FirstOrDefault(s => s.ColIndex == attribute.ColIndex).Value;
                    try
                    {
                        SetProperty(result, type.GetProperty(item.Name), value);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"invalid value at col {attribute.ColIndex}.detail:{ex.Message}");
                    }
                }
            }

            return result;
        }

        private static void SetProperty<T>(T result, PropertyInfo propertyInfo, object value) where T : new()
        {
            var type = propertyInfo.PropertyType;
            MethodInfo setter = propertyInfo.GetSetMethod();
            //Int
            if (type.Equals(typeof(Int32)))
            {
                var PropSet = (PropertySetter<int>)Delegate.CreateDelegate(typeof(PropertySetter<int>), result, setter);
                PropSet(Convert.ToInt32(value));
            }
            else if (type.Equals(typeof(Int32?)))
            {
                if (value != null && !string.IsNullOrEmpty(Convert.ToString(value))) 
                {
                    var PropSet = (PropertySetter<int?>)Delegate.CreateDelegate(typeof(PropertySetter<int?>), result, setter);
                    PropSet(Convert.ToInt32(value));
                }
            }
            //Decimal
            else if (type.Equals(typeof(decimal)))
            {
                var PropSet = (PropertySetter<decimal>)Delegate.CreateDelegate(typeof(PropertySetter<decimal>), result, setter);
                PropSet(Convert.ToDecimal(value));
            }
            else if (type.Equals(typeof(decimal?)))
            {
                if (value != null && !string.IsNullOrEmpty(Convert.ToString(value)))
                {
                    var PropSet = (PropertySetter<decimal?>)Delegate.CreateDelegate(typeof(PropertySetter<decimal?>), result, setter);
                    PropSet(Convert.ToDecimal(value));
                }
            }
            //Bool
            else if (type.Equals(typeof(bool)))
            {
                var PropSet = (PropertySetter<bool>)Delegate.CreateDelegate(typeof(PropertySetter<bool>), result, setter);
                PropSet(Convert.ToBoolean(value));
            }
            else if (type.Equals(typeof(bool?)))
            {
                if (value != null)
                {
                    var PropSet = (PropertySetter<bool?>)Delegate.CreateDelegate(typeof(PropertySetter<bool?>), result, setter);
                    PropSet(Convert.ToBoolean(value));
                }
            }
            //String
            else if (type.Equals(typeof(String)))
            {
                var PropSet = (PropertySetter<string>)Delegate.CreateDelegate(typeof(PropertySetter<string>), result, setter);
                PropSet(Convert.ToString(value));
            }
            //DateTime
            else if (type.Equals(typeof(DateTime)))
            {
                if (value != null)
                {
                    var PropSet = (PropertySetter<DateTime>)Delegate.CreateDelegate(typeof(PropertySetter<DateTime>), result, setter);
                    PropSet(Convert.ToDateTime(value));
                }
            }
            else if (type.Equals(typeof(DateTime?)))
            {
                if (value != null)
                {
                    var PropSet = (PropertySetter<DateTime?>)Delegate.CreateDelegate(typeof(PropertySetter<DateTime?>), result, setter);
                    PropSet(Convert.ToDateTime(value));
                }
            }
            //byte[]
            else if (type.Equals(typeof(byte[])))
            {
                if (value != null)
                {
                    var PropSet = (PropertySetter<byte[]>)Delegate.CreateDelegate(typeof(PropertySetter<byte[]>), result, setter);
                    PropSet((byte[])value);
                }
            }
            else
            {
                //无法识别的属性，不能使用泛型委托
                propertyInfo.SetValue(result, value, null);
            }
        }
    }
}
