using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Scb
{
    /// <summary>
    /// Provides access to commonly used reflection features via a relatively simple interface.  
    /// Also, instead of returning null for an unknown member (as the reflection methods generally
    /// do), this class propagates an exception with a useful message.
    /// </summary>
    public static class ReflectionUtil
    {
        private const BindingFlags InstanceFlags = BindingFlags.Public | BindingFlags.Instance;
        private const BindingFlags StaticFlags = BindingFlags.Public | BindingFlags.Static;

        private static string Describe(BindingFlags bindingFlags)
        {
            var names = new List<string>();
            foreach (BindingFlags bindingFlag in Enum.GetValues(typeof(BindingFlags)))
                if (bindingFlags.HasFlag(bindingFlag))
                    names.Add(bindingFlag.ToString());
            return "[" + string.Join("|", names) + "]";
        }

        /// <summary>
        /// Returns the field with the specified name of a type or propagates 
        /// ArgumentException if no match.
        /// </summary>
        private static FieldInfo GetField(Type type, string fieldName, BindingFlags bindingFlags)
        {
            var field = type.GetField(fieldName, bindingFlags);
            if (field == null)
                throw new ArgumentException("Type '" + type.Name + "' has no field named '" + fieldName + "' for binding " + Describe(bindingFlags) + ".");
            return field;
        }

        /// <summary>
        /// Returns the property with the specified name of a type or propagates 
        /// ArgumentException if no match.
        /// </summary>
        private static PropertyInfo GetProperty(Type type, string propertyName, BindingFlags bindingFlags)
        {
            var property = type.GetProperty(propertyName, bindingFlags);
            if (property == null)
                throw new ArgumentException("Type '" + type.Name + "' has no property named '" + propertyName + "' for binding " + Describe(bindingFlags) + ".");
            return property;
        }

        /// <summary>
        /// Returns the field or property with the specified name of a type or propagates 
        /// ArgumentException if no match.
        /// </summary>
        private static MemberInfo GetFieldOrProperty(Type type, string fieldOrPropertyName, BindingFlags bindingFlags)
        {
            var field = type.GetField(fieldOrPropertyName, bindingFlags);
            if (field != null)
                return field;
            var property = type.GetProperty(fieldOrPropertyName, bindingFlags);
            if (property != null)
                return property;
            throw new ArgumentException("Type '" + type.Name + "' has no field or property named '" + fieldOrPropertyName + "' for binding " + Describe(bindingFlags) + ".");
        }

        /// <summary>
        /// Returns the indexer property of a type or propagates ArgumentException if the type has 
        /// no indexer.
        /// </summary>
        private static PropertyInfo GetIndexer(Type type)
        {
            foreach (var property in type.GetProperties())
                if (property.GetIndexParameters().Length > 0)
                    return property;
            throw new ArgumentException("Type '" + type.Name + "' has no indexer property.");
        }

        /// <summary>
        /// Returns the method with the specified name or propagates ArgumentException if no match.  
        /// A system exception is propagated if the method is overloaded.
        /// </summary>
        private static MethodInfo GetMethod(Type type, string methodName, BindingFlags bindingFlags)
        {
            var method = type.GetMethod(methodName, bindingFlags);
            if (method == null)
                throw new ArgumentException("Method '" + type.Name + "." + methodName + "' not found for binding " + Describe(bindingFlags) + ".");
            return method;
        }

        /// <summary>
        /// Returns the method with the specified name and parameter types or propagates 
        /// ArgumentException if no match.
        /// </summary>
        private static MethodInfo GetMethod(Type type, string methodName, Type[] types, BindingFlags bindingFlags)
        {
            var method = type.GetMethod(methodName, bindingFlags, null, types, null);
            if (method == null)
            {
                string methodSignature = methodName + "(" + string.Join(",", types.Select(t => t.Name)) + ")";
                throw new ArgumentException("Method '" + type.Name + "." + methodSignature + "' not found for binding " + Describe(bindingFlags) + ".");
            }
            return method;
        }

        /// <summary>
        /// Returns the value of a static field/property with the specified name.
        /// </summary>
        private static object Read(Type type, string fieldOrPropertyName)
        {
            var fieldOrProperty = GetFieldOrProperty(type, fieldOrPropertyName, StaticFlags);
            var field = fieldOrProperty as FieldInfo;
            if (field != null)
                return field.GetValue(null);
            return ((PropertyInfo)fieldOrProperty).GetValue(null);
        }

        /// <summary>
        /// Sets the value of a static field/property with the specified name.
        /// </summary>
        private static void Write(Type type, string fieldOrPropertyName, object value)
        {
            var fieldOrProperty = GetFieldOrProperty(type, fieldOrPropertyName, StaticFlags);
            var field = fieldOrProperty as FieldInfo;
            if (field != null)
                field.SetValue(null, value);
            else
                ((PropertyInfo)fieldOrProperty).SetValue(null, value);
        }

        /// <summary>
        /// Invokes the method with the specified name and signature defined by the type of each 
        /// parameter.
        /// </summary>
        private static object InvokeMethodWithNonNullParameters(Type type, object instance, string methodName, BindingFlags bindingFlags, object[] parameters)
        {
            if (parameters.Any(p => p == null))
                throw new NullReferenceException("All parameters must be non-null.");
            var types = parameters.Select(p => p.GetType()).ToArray();
            MethodInfo methodInfo = GetMethod(type, methodName, types, bindingFlags);
            return methodInfo.Invoke(instance, parameters);
        }

        #region Type search

        /// <summary>
        /// Searches all loaded assemblies for a type with the specified name.  If the name is 
        /// unique, then returns the type.  Propagates ArgumentException if either no matches or 
        /// more than one.
        /// </summary>
        public static Type GetExpectedType(string typeName)
        {
            Type foundType = null;
            foreach (var assembly in AppDomain.CurrentDomain.GetAssemblies())
            {
                Type type = assembly.GetType(typeName);
                if (type != null)
                {
                    if (foundType != null)
                        throw new ArgumentException("More than one type named '" + typeName + "' in loaded assemblies.");
                    foundType = type;
                }
            }
            if (foundType == null)
                throw new ArgumentException("Type '" + typeName + "' not found in any loaded assembly.");
            return foundType;
        }

        /// <summary>
        /// Returns the type with the specified name of an assembly or propagates 
        /// ArgumentException if no match.
        /// </summary>
        public static Type GetExpectedType(this Assembly assembly, string typeName)
        {
            Type type = assembly.GetType(typeName);
            if (type == null)
                throw new ArgumentException("Type '" + typeName + "' not found in assembly '" + assembly.Location + "'.");
            return type;
        }

        #endregion

        #region Type member search

        /// <summary>
        /// Returns the field with the specified name of a type or propagates 
        /// ArgumentException if no match.
        /// </summary>
        public static FieldInfo GetExpectedField(this Type type, string fieldName, BindingFlags bindingFlags = InstanceFlags)
        {
            return GetField(type, fieldName, bindingFlags);
        }

        /// <summary>
        /// Returns the property with the specified name of a type or propagates 
        /// ArgumentException if no match.
        /// </summary>
        public static PropertyInfo GetExpectedProperty(this Type type, string propertyName, BindingFlags bindingFlags = InstanceFlags)
        {
            return GetProperty(type, propertyName, bindingFlags);
        }

        /// <summary>
        /// Returns the field or property with the specified name of a type or propagates 
        /// ArgumentException if no match.
        /// </summary>
        public static MemberInfo GetExpectedFieldOrProperty(this Type type, string fieldOrPropertyName, BindingFlags bindingFlags = InstanceFlags)
        {
            return GetFieldOrProperty(type, fieldOrPropertyName, bindingFlags);
        }

        /// <summary>
        /// Returns the indexer property of a type or propagates ArgumentException if the type has 
        /// no indexer.
        /// </summary>
        public static PropertyInfo GetExpectedIndexer(this Type type)
        {
            return GetIndexer(type);
        }

        /// <summary>
        /// Returns the method with the specified name or propagates ArgumentException if no match.  
        /// A system exception is propagated if the method is overloaded.
        /// </summary>
        public static MethodInfo GetExpectedMethod(this Type type, string methodName, BindingFlags bindingFlags = InstanceFlags)
        {
            return GetMethod(type, methodName, bindingFlags);
        }

        /// <summary>
        /// Returns the method with the specified name and parameter types or propagates 
        /// ArgumentException if no match.
        /// </summary>
        public static MethodInfo GetExpectedMethod(this Type type, string methodName, Type[] types, BindingFlags bindingFlags = InstanceFlags)
        {
            return GetMethod(type, methodName, types, bindingFlags);
        }

        #endregion

        #region Static (class) access

        /// <summary>
        /// Returns the value of the static field/property with the specified name.
        /// </summary>
        public static object ReadStatic(this Type type, string fieldOrPropertyName)
        {
            return Read(type, fieldOrPropertyName);
        }

        /// <summary>
        /// Sets the value of the static field/property with the specified name.
        /// </summary>
        public static void WriteStatic(this Type type, string fieldOrPropertyName, object value)
        {
            Write(type, fieldOrPropertyName, value);
        }

        /// <summary>
        /// Invokes the static method with the specified name.  Fails if method is overloaded.
        /// </summary>
        public static object InvokeStaticMethod(this Type type, string methodName, params object[] parameters)
        {
            var method = GetMethod(type, methodName, StaticFlags);
            return method.Invoke(null, parameters);
        }

        /// <summary>
        /// Invokes the static method with the specified name and type signature specified by the 
        /// parameter types.  Fails if any parameter is null.
        /// </summary>
        public static object InvokeStaticMethodWithNonNullParameters(this Type type, string methodName, params object[] parameters)
        {
            return InvokeMethodWithNonNullParameters(type, (object)null, methodName, StaticFlags, parameters);
        }

        /// <summary>
        /// Invokes the static method with the specified name and type signature.
        /// </summary>
        public static object InvokeStaticMethod(this Type type, string methodName, Type[] types, object[] parameters)
        {
            var method = GetMethod(type, methodName, types, StaticFlags);
            return method.Invoke(null, parameters);
        }

        #endregion

        #region Object (instance) access

        /// <summary>
        /// Returns the value of the field/property with the specified name.
        /// </summary>
        public static object ReadViaReflection(this object instance, string fieldOrPropertyName)
        {
            Type type = instance.GetType();
            MemberInfo fieldOrProperty = GetFieldOrProperty(type, fieldOrPropertyName, InstanceFlags);
            var field = fieldOrProperty as FieldInfo;
            if (field != null)
                return field.GetValue(instance);
            return ((PropertyInfo)fieldOrProperty).GetValue(instance);
        }

        /// <summary>
        /// Sets the value of the field/property with the specified name.
        /// </summary>
        public static void WriteViaReflection(this object instance, string fieldOrPropertyName, object value)
        {
            var type = instance.GetType();
            MemberInfo fieldOrProperty = GetFieldOrProperty(type, fieldOrPropertyName, InstanceFlags);
            var field = fieldOrProperty as FieldInfo;
            if (field != null)
                field.SetValue(instance, value);
            else
                ((PropertyInfo)fieldOrProperty).SetValue(instance, value);
        }

        /// <summary>
        /// Returns the value of an object's indexer property.
        /// </summary>
        public static object ReadIndexerViaReflection(this object instance, params object[] indexes)
        {
            var type = instance.GetType();
            PropertyInfo indexer = GetIndexer(type);
            return indexer.GetValue(instance, indexes);
        }

        /// <summary>
        /// Sets the value of an object's indexer property.
        /// </summary>
        public static void WriteIndexerViaReflection(this object instance, object value, params object[] indexes)
        {
            var type = instance.GetType();
            PropertyInfo indexer = GetIndexer(type);
            indexer.SetValue(instance, value, indexes);
        }

        /// <summary>
        /// Invokes the method with the specified name and type signature.
        /// </summary>
        /// <remarks>
        /// This is the most precise and flexible variant for invoking a method, but requires more 
        /// effort to use since the types array must match the desired method.
        /// </remarks>
        public static object InvokeMethodViaReflection(this object instance, string methodName, Type[] types, object[] parameters)
        {
            var type = instance.GetType();
            var method = GetMethod(type, methodName, types, InstanceFlags);
            return method.Invoke(instance, parameters);
        }

        /// <summary>
        /// Invokes the method with the specified name.  Fails if method is overloaded.
        /// </summary>
        /// <remarks>
        /// This is the easiest variant for invoking a method since it requires the least in terms 
        /// of type information, but cannot be used if the method is overloaded with same number 
        /// of parameters.
        /// </remarks>
        public static object InvokeMethodViaReflection(this object instance, string methodName, params object[] parameters)
        {
            var type = instance.GetType();
            var method = GetMethod(type, methodName, InstanceFlags);
            return method.Invoke(instance, parameters);
        }

        /// <summary>
        /// Invokes the method with the specified name and type signature specified by the 
        /// parameter types.  Fails if any parameter is null.
        /// </summary>
        /// <remarks>
        /// This variant for invoking a method allows selecting between overloaded methods with the
        /// same number of parameters, but only if each parameter value is non-null.
        /// </remarks>
        public static object InvokeMethodWithNonNullParametersViaReflection(this object instance, string methodName, params object[] parameters)
        {
            var type = instance.GetType();
            return InvokeMethodWithNonNullParameters(type, instance, methodName, InstanceFlags, parameters);
        }

        #endregion
    }
}