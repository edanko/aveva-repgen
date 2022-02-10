using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.ApplicationServices;
using Microsoft.VisualBasic.CompilerServices;

namespace ReportsGenerator.My;

[HideModuleName]
[StandardModule]
[GeneratedCode("MyTemplate", "10.0.0.0")]
internal sealed class MyProject
{
    private static readonly ThreadSafeObjectProvider<MyComputer> MComputerObjectProvider = new();

    private static readonly ThreadSafeObjectProvider<MyApplication> MAppObjectProvider = new();

    private static readonly ThreadSafeObjectProvider<User> MUserObjectProvider = new();

    private static readonly ThreadSafeObjectProvider<MyForms> _mMyFormsObjectProvider = new();

    [HelpKeyword("My.Computer")]
    internal static MyComputer Computer
    {
        [DebuggerHidden] get => MComputerObjectProvider.GetInstance;
    }

    [HelpKeyword("My.Application")]
    internal static MyApplication Application
    {
        [DebuggerHidden] get => MAppObjectProvider.GetInstance;
    }

    [HelpKeyword("My.User")]
    internal static User User
    {
        [DebuggerHidden] get => MUserObjectProvider.GetInstance;
    }

    [HelpKeyword("My.Forms")]
    internal static MyForms Forms
    {
        [DebuggerHidden] get => _mMyFormsObjectProvider.GetInstance;
    }

    [MyGroupCollection("System.Windows.Forms.Form", "Create__Instance__", "Dispose__Instance__", "My.MyProject.Forms")]
    [EditorBrowsable(EditorBrowsableState.Never)]
    internal sealed class MyForms
    {
        [ThreadStatic] private static Hashtable _mFormBeingCreated;

        public Form2 MForm1;

        [EditorBrowsable(EditorBrowsableState.Never)]
        [DebuggerHidden]
        public MyForms()
        {
        }

        public Form2 Form1
        {
            get
            {
                MForm1 = Create__Instance__(MForm1);
                return MForm1;
            }
            set
            {
                if (value == MForm1) return;
                if (value != null) throw new ArgumentException("Property can only be set to Nothing");
                Dispose__Instance__(ref MForm1);
            }
        }

        [DebuggerHidden]
        private static T Create__Instance__<T>(T instance) where T : Form, new()
        {
            if (instance == null || instance.IsDisposed)
            {
                if (_mFormBeingCreated != null)
                {
                    if (_mFormBeingCreated.ContainsKey(typeof(T)))
                        throw new InvalidOperationException(Utils.GetResourceString("WinForms_RecursiveFormCreate"));
                }
                else
                {
                    _mFormBeingCreated = new Hashtable();
                }

                _mFormBeingCreated.Add(typeof(T), null);
                try
                {
                    return Activator.CreateInstance<T>();
                }
                catch (TargetInvocationException ex) when (ex.InnerException != null)
                {
                    var resourceString =
                        Utils.GetResourceString("WinForms_SeeInnerException", ex.InnerException.Message);
                    throw new InvalidOperationException(resourceString, ex.InnerException);
                }
                finally
                {
                    _mFormBeingCreated.Remove(typeof(T));
                }
            }

            return instance;
        }

        [DebuggerHidden]
        private void Dispose__Instance__<T>(ref T instance) where T : Form
        {
            instance.Dispose();
            instance = default;
        }

        [EditorBrowsable(EditorBrowsableState.Never)]
        public override bool Equals(object o)
        {
            return base.Equals(o);
        }

        [EditorBrowsable(EditorBrowsableState.Never)]
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        [EditorBrowsable(EditorBrowsableState.Never)]
        internal new Type GetType()
        {
            return typeof(MyForms);
        }

        [EditorBrowsable(EditorBrowsableState.Never)]
        public override string ToString()
        {
            return base.ToString();
        }
    }


    [EditorBrowsable(EditorBrowsableState.Never)]
    [ComVisible(false)]
    internal sealed class ThreadSafeObjectProvider<T> where T : new()
    {
        [ThreadStatic] [CompilerGenerated] private static T _mThreadStaticValue;

        [EditorBrowsable(EditorBrowsableState.Never)]
        [DebuggerHidden]
        public ThreadSafeObjectProvider()
        {
        }

        internal T GetInstance
        {
            [DebuggerHidden]
            get
            {
                if (_mThreadStaticValue == null) _mThreadStaticValue = Activator.CreateInstance<T>();
                return _mThreadStaticValue;
            }
        }
    }
}