using System;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Scb
{
    internal class TestReflectionClass0 { }

    internal class TestReflectionClass1
    {
        public int PublicField;
        public int PublicProperty { get; set; }
        public static int StaticPublicField;
        public static int StaticPublicProperty { get; set; }

        public int[] IndexedValue;
        public int this[int i]
        {
            get { return IndexedValue[i]; }
            set { IndexedValue[i] = value; }
        }

        public bool PublicParameterlessMethodCalled;
        public void PublicParameterlessMethod()
        {
            PublicParameterlessMethodCalled = true;
        }

        public int PublicMethodWithIntParameterData;
        public int PublicMethodWithIntParameter(int data)
        {
            PublicMethodWithIntParameterData = data;
            return data;
        }

        public static bool StaticPublicParameterlessMethodCalled;
        public static void StaticPublicParameterlessMethod()
        {
            StaticPublicParameterlessMethodCalled = true;
        }

        public static string StaticPublicMethodWithStringParameterData;
        public static string StaticPublicMethodWithStringParameter(string data)
        {
            StaticPublicMethodWithStringParameterData = data;
            return data;
        }

        public void PublicOverloadedMethod() { }
        public void PublicOverloadedMethod(int i) { }

        public static void PublicStaticOverloadedMethod() { }
        public static void PublicStaticOverloadedMethod(int i) { }
    }
    
    [TestClass]
    public class ReflectionUtilUnitTest
    {
        #region GetExpectedType

        [TestMethod]
        public void GetExpectedType_IsType()
        {
            Assert.IsNotNull(ReflectionUtil.GetExpectedType(typeof(TestReflectionClass1).FullName));
        }

        [TestMethod]
        public void GetExpectedType_Propagates_ForUnkownType()
        {
            ExceptionAssert.Propagates<ArgumentException>(() => ReflectionUtil.GetExpectedType("NOTTHERE"));
        }

        //TODO: Add test to verify exception if find more than one type with same name in loaded assemblies ... but how??

        [TestMethod]
        public void GetExpectedType_IsTypeOfAssembly()
        {
            Assembly executingAssembly = Assembly.GetExecutingAssembly();
            Assert.IsNotNull(executingAssembly.GetExpectedType(typeof(TestReflectionClass1).FullName));
        }

        [TestMethod]
        public void GetExpectedType_Propagates_ForUnkownTypeOfAssembly()
        {
            Assembly executingAssembly = Assembly.GetExecutingAssembly();
            ExceptionAssert.Propagates<ArgumentException>(() => executingAssembly.GetExpectedType("NOTTHERE"));
        }

        #endregion

        #region GetExpectedField

        [TestMethod]
        public void GetExpectedField_IsFieldInfo()
        {
            Assert.IsNotNull(typeof(TestReflectionClass1).GetExpectedField("PublicField"));
        }

        [TestMethod]
        public void GetExpectedField_Propagates_ForUnknownFieldName()
        {
            ExceptionAssert.Propagates<ArgumentException>(() => Assert.IsNull(typeof(TestReflectionClass1).GetExpectedField("NOTTHERE")));
        }

        #endregion

        #region GetExpectedProperty

        [TestMethod]
        public void GetExpectedProperty_IsPropertyInfo()
        {
            Assert.IsNotNull(typeof(TestReflectionClass1).GetExpectedProperty("PublicProperty"));
        }

        [TestMethod]
        public void GetExpectedProperty_Propagates_ForUnknownPropertyName()
        {
            ExceptionAssert.Propagates<ArgumentException>(() => typeof(TestReflectionClass1).GetExpectedProperty("NOTTHERE"));
        }

        #endregion

        #region GetExpectedFieldOrProperty

        [TestMethod]
        public void GetExpectedFieldOrProperty_IsFieldInfo()
        {
            Assert.IsNotNull(typeof(TestReflectionClass1).GetExpectedFieldOrProperty("PublicField"));
        }

        [TestMethod]
        public void GetExpectedFieldOrProperty_IsPropertyInfo()
        {
            Assert.IsNotNull(typeof(TestReflectionClass1).GetExpectedFieldOrProperty("PublicProperty"));
        }

        [TestMethod]
        public void GetExpectedFieldOrProperty_Propagates_ForUnknownFieldOrPropertyName()
        {
            ExceptionAssert.Propagates<ArgumentException>(() => typeof(TestReflectionClass1).GetExpectedFieldOrProperty("NOTTHERE"));
        }

        #endregion

        #region GetExpectedIndexer

        [TestMethod]
        public void GetExpectedIndexer_IsPropertyInfo()
        {
            Assert.IsNotNull(typeof(TestReflectionClass1).GetExpectedIndexer());
        }

        [TestMethod]
        public void GetExpectedIndexer_Propagates_ForClassWithNoIndexer()
        {
            ExceptionAssert.Propagates<ArgumentException>(() => typeof(TestReflectionClass0).GetExpectedIndexer());
        }

        #endregion

        #region GetExpectedMethod(string, BindingFlags)

        [TestMethod]
        public void GetExpectedMethod_IsMethodInfo()
        {
            Assert.IsNotNull(typeof(TestReflectionClass1).GetExpectedMethod("PublicParameterlessMethod"));
        }

        [TestMethod]
        public void GetExpectedMethod_Propagates_ForUnknownMethodName()
        {
            ExceptionAssert.Propagates<ArgumentException>(() => typeof(TestReflectionClass1).GetExpectedMethod("NOTTHERE"));
        }

        #endregion

        #region GetExpectedMethod(string, type[], BindingFlags)

        [TestMethod]
        public void GetExpectedMethodMatchingTypes_IsMethodInfo()
        {
            Assert.AreEqual(0, typeof(TestReflectionClass1).GetExpectedMethod("PublicOverloadedMethod", new Type[0]).GetParameters().Length);
            Assert.AreEqual(1, typeof(TestReflectionClass1).GetExpectedMethod("PublicOverloadedMethod", new[] { typeof(int) }).GetParameters().Length);
        }

        [TestMethod]
        public void GetExpectedMethodMatchingTypes_Propagates_ForUnknownMethodName()
        {
            ExceptionAssert.Propagates<ArgumentException>(() => typeof(TestReflectionClass1).GetExpectedMethod("PublicOverloadedMethod", new[] { typeof(string) }));
        }

        #endregion

        #region Read/write static field/property

        [TestMethod]
        public void ReadStatic_IsStaticFieldValue()
        {
            TestReflectionClass1.StaticPublicField = 5;
            Assert.AreEqual(5, typeof(TestReflectionClass1).ReadStatic("StaticPublicField"));
        }

        [TestMethod]
        public void ReadStatic_IsStaticPropertyValue()
        {
            TestReflectionClass1.StaticPublicProperty = 6;
            Assert.AreEqual(6, typeof(TestReflectionClass1).ReadStatic("StaticPublicProperty"));
        }

        [TestMethod]
        public void WriteStatic_SetsStaticFieldValue()
        {
            typeof(TestReflectionClass1).WriteStatic("StaticPublicField", 18);
            Assert.AreEqual(18, TestReflectionClass1.StaticPublicField);
        }

        [TestMethod]
        public void WriteStatic_SetsStaticPropertyValue()
        {
            typeof(TestReflectionClass1).WriteStatic("StaticPublicProperty", 28);
            Assert.AreEqual(28, TestReflectionClass1.StaticPublicProperty);
        }

        #endregion

        #region Invoke static method with with matching name and type signature

        [TestMethod]
        public void InvokeStaticMethodBySignature_CallsStaticMethod()
        {
            TestReflectionClass1.StaticPublicParameterlessMethodCalled = false;
            typeof(TestReflectionClass1).InvokeStaticMethod("StaticPublicParameterlessMethod", new Type[0], new object[0]);
            Assert.IsTrue(TestReflectionClass1.StaticPublicParameterlessMethodCalled);
        }

        [TestMethod]
        public void InvokeStaticMethodBySignature_PassesParameters()
        {
            TestReflectionClass1.StaticPublicMethodWithStringParameterData = "xxx";
            typeof(TestReflectionClass1).InvokeStaticMethod("StaticPublicMethodWithStringParameter",
                                                             new[] { typeof(string) }, new object[] { "abc" });
            Assert.AreEqual("abc", TestReflectionClass1.StaticPublicMethodWithStringParameterData);
        }

        [TestMethod]
        public void InvokeStaticMethodBySignature_IsResult()
        {
            var result = typeof(TestReflectionClass1).InvokeStaticMethod("StaticPublicMethodWithStringParameter",
                                                             new[] { typeof(string) }, new object[] { "xyz" });
            Assert.AreEqual("xyz", result);
        }

        #endregion

        #region Invoke static method with matching name only

        [TestMethod]
        public void InvokeStaticMethodByName_CallsStaticMethod()
        {
            TestReflectionClass1.StaticPublicParameterlessMethodCalled = false;
            typeof(TestReflectionClass1).InvokeStaticMethod("StaticPublicParameterlessMethod");
            Assert.IsTrue(TestReflectionClass1.StaticPublicParameterlessMethodCalled);
        }

        [TestMethod]
        public void InvokeStaticMethodByName_PassesParameters()
        {
            TestReflectionClass1.StaticPublicMethodWithStringParameterData = "xxx";
            typeof(TestReflectionClass1).InvokeStaticMethod("StaticPublicMethodWithStringParameter", "abc");
            Assert.AreEqual("abc", TestReflectionClass1.StaticPublicMethodWithStringParameterData);
        }

        [TestMethod]
        public void InvokeStaticMethodByName_IsResult()
        {
            var result = typeof(TestReflectionClass1).InvokeStaticMethod("StaticPublicMethodWithStringParameter", "xyz");
            Assert.AreEqual("xyz", result);
        }

        #endregion

        #region Invoke static method with non-null parameters

        [TestMethod]
        public void InvokeStaticMethodWithNonNullParameters_CallsStaticMethod()
        {
            TestReflectionClass1.StaticPublicParameterlessMethodCalled = false;
            typeof(TestReflectionClass1).InvokeStaticMethodWithNonNullParameters("StaticPublicParameterlessMethod");
            Assert.IsTrue(TestReflectionClass1.StaticPublicParameterlessMethodCalled);
        }

        [TestMethod]
        public void InvokeStaticMethodWithNonNullParameters_PassesParameters()
        {
            TestReflectionClass1.StaticPublicMethodWithStringParameterData = "xxx";
            typeof(TestReflectionClass1).InvokeStaticMethodWithNonNullParameters("StaticPublicMethodWithStringParameter", "abc");
            Assert.AreEqual("abc", TestReflectionClass1.StaticPublicMethodWithStringParameterData);
        }

        [TestMethod]
        public void InvokeStaticMethodWithNonNullParameters_PropagatesNullReferenceExceptionNullParameter()
        {
            ExceptionAssert.Propagates<NullReferenceException>(() =>
                typeof(TestReflectionClass1).InvokeStaticMethodWithNonNullParameters("StaticPublicMethodWithStringParameter", new object[] { null }));
        }

        [TestMethod]
        public void InvokeStaticMethodWithNonNullParameters_IsResult()
        {
            var result = typeof(TestReflectionClass1).InvokeStaticMethodWithNonNullParameters("StaticPublicMethodWithStringParameter", "xyz");
            Assert.AreEqual("xyz", result);
        }

        #endregion
        
        #region Read/write field/property

        [TestMethod]
        public void ReadViaReflection_IsFieldValue()
        {
            var instance = new TestReflectionClass1 { PublicField = 5 };
            Assert.AreEqual(5, instance.ReadViaReflection("PublicField"));
        }

        [TestMethod]
        public void ReadViaReflection_IsPropertyValue()
        {
            var instance = new TestReflectionClass1 { PublicProperty = 6 };
            Assert.AreEqual(6, instance.ReadViaReflection("PublicProperty"));
        }

        [TestMethod]
        public void WriteViaReflection_SetsFieldValue()
        {
            var instance = new TestReflectionClass1();
            instance.WriteViaReflection("PublicField", 18);
            Assert.AreEqual(18, instance.PublicField);
        }

        [TestMethod]
        public void WriteViaReflection_SetsPropertyValue()
        {
            var instance = new TestReflectionClass1();
            instance.WriteViaReflection("PublicProperty", 28);
            Assert.AreEqual(28, instance.PublicProperty);
        }

        #endregion

        #region Read/write indexer property

        [TestMethod]
        public void ReadIndexeViaReflectionr_IsIndexerValue()
        {
            var instance = new TestReflectionClass1();
            instance.IndexedValue = new int[1];
            instance[0] = 101;
            Assert.AreEqual(101, instance.ReadIndexerViaReflection(0));
        }

        [TestMethod]
        public void WriteIndexerViaReflection_SetsIndexerValue()
        {
            var instance = new TestReflectionClass1();
            instance.IndexedValue = new int[1];
            instance.WriteIndexerViaReflection(8, 0);
            Assert.AreEqual(8, instance[0]);
        }

        #endregion

        #region Invoke method with with matching name and type signature

        [TestMethod]
        public void InvokeMethodViaReflection_WithTypes_CallsMethod()
        {
            var instance = new TestReflectionClass1();
            instance.InvokeMethodViaReflection("PublicParameterlessMethod", new Type[0], new object[0]);
            Assert.IsTrue(instance.PublicParameterlessMethodCalled);
        }

        [TestMethod]
        public void InvokeMethodViaReflection_WithTypes_PassesParameters()
        {
            var instance = new TestReflectionClass1();
            instance.InvokeMethodViaReflection("PublicMethodWithIntParameter", new[] { typeof(int) }, new object[] { 101 });
            Assert.AreEqual(101, instance.PublicMethodWithIntParameterData);
        }

        [TestMethod]
        public void InvokeMethodViaReflection_WithTypes_IsResult()
        {
            var instance = new TestReflectionClass1();
            var result = instance.InvokeMethodViaReflection("PublicMethodWithIntParameter", new[] { typeof(int) }, new object[] { 45 });
            Assert.AreEqual(45, result);
        }

        #endregion

        #region Invoke method with matching name only

        [TestMethod]
        public void InvokeMethodViaReflection_WithoutTypes_CallsMethod()
        {
            var instance = new TestReflectionClass1();
            instance.InvokeMethodViaReflection("PublicParameterlessMethod");
            Assert.IsTrue(instance.PublicParameterlessMethodCalled);
        }

        [TestMethod]
        public void InvokeMethodViaReflection_WithoutTypes_PassesParameters()
        {
            var instance = new TestReflectionClass1();
            instance.InvokeMethodViaReflection("PublicMethodWithIntParameter", 66);
            Assert.AreEqual(66, instance.PublicMethodWithIntParameterData);
        }

        [TestMethod]
        public void InvokeMethodViaReflection_WithoutTypes_IsResult()
        {
            var instance = new TestReflectionClass1();
            var result = instance.InvokeMethodViaReflection("PublicMethodWithIntParameter", 77);
            Assert.AreEqual(77, result);
        }

        #endregion

        #region Invoke method with non-null parameters

        [TestMethod]
        public void InvokeMethodWithNonNullParametersViaReflection_CallsMethod()
        {
            var instance = new TestReflectionClass1();
            instance.InvokeMethodWithNonNullParametersViaReflection("PublicParameterlessMethod");
            Assert.IsTrue(instance.PublicParameterlessMethodCalled);
        }

        [TestMethod]
        public void InvokeMethodWithNonNullParametersViaReflection_PassesParameters()
        {
            var instance = new TestReflectionClass1();
            instance.InvokeMethodWithNonNullParametersViaReflection("PublicMethodWithIntParameter",88);
            Assert.AreEqual(88, instance.PublicMethodWithIntParameterData);
        }

        [TestMethod]
        public void InvokeMethodWithNonNullParametersViaReflection_PropagatesNullReferenceExceptionNullParameter()
        {
            var instance = new TestReflectionClass1();
            ExceptionAssert.Propagates<NullReferenceException>(() =>
                instance.InvokeMethodWithNonNullParametersViaReflection("PublicMethodWithIntParameter", new object[] { null }));
        }

        [TestMethod]
        public void InvokeMethodWithNonNullParametersViaReflection_IsResult()
        {
            var instance = new TestReflectionClass1();
            var result = instance.InvokeMethodWithNonNullParametersViaReflection("PublicMethodWithIntParameter", 99);
            Assert.AreEqual(99, result);
        }

        #endregion
    }
}
