using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.ComponentModel;

namespace NetOffice.Tools.Native.Bridge
{
    /// <summary>
    /// Represents a handle to an unmanaged library
    /// </summary>
    [DebuggerDisplay("{Name}")]
    public class CdeclHandle : IDisposable
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="ptr">underlying handle ptr</param>
        /// <param name="name">name of the library</param>
        public CdeclHandle(IntPtr ptr, string name)
        {
            Underlying = ptr;
            Name = name;
            Functions = new Dictionary<string, Delegate>();
        }

        /// <summary>
        /// Underyling Handle Ptr is empty
        /// </summary>
        public bool HandleIsZero
        {
            get
            {
                return Underlying != IntPtr.Zero;
            }
        }

        /// <summary>
        /// Name of the Library
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Underlying Library Handle
        /// </summary>
        private IntPtr Underlying { get; set; }

        /// <summary>
        /// Delegate Cache
        /// </summary>
        private Dictionary<string, Delegate> Functions { get; set; }

        /// <summary>
        /// Returns a function pointer by name. The method is caching the operation.
        /// </summary>
        /// <param name="name">name of the method</param>
        /// <param name="type">target delegate type</param>
        /// <returns>delegate to unmanaged method</returns>
        /// <exception cref="Win32Exception">Unable to get proc address or function pointer</exception>
        /// <exception cref="ArgumentNullException">an argument is null or empty</exception>
        public Delegate GetDelegateForFunctionPointer(string name, Type type)
        {
            if (String.IsNullOrWhiteSpace(name))
                throw new ArgumentNullException("name");
            if (null == type)
                throw new ArgumentNullException("type");

            Delegate result = null;
            if (!Functions.ContainsKey(name))
            {
                IntPtr ptr = Interop.GetProcAddress(Underlying, name);
                if (ptr == IntPtr.Zero)
                    throw new Win32Exception(String.Format("Unable to get proc address <{0}> in <{1}>.", name, Name));
                result = Marshal.GetDelegateForFunctionPointer(ptr, type) as Delegate;
                if (null == result)
                    throw new Win32Exception(String.Format("Unable to get function pointer <{0}> in <{1}>.", name, Name));              
                Functions.Add(name, result);
                return result;
            }
            else
                result = Functions[name];

            return result;
        }

        /// <summary>
        /// Loads an unmanaged library from filesystem
        /// </summary>
        /// <param name="fullFileName">full qualified name of the library file</param>
        /// <param name="version">optional file version to check</param>
        /// <returns>handle to library</returns>
        /// <exception cref="FileNotFoundException">File is missing</exception>
        /// <exception cref="Win32Exception">Unable to load library</exception>
        /// <exception cref="FileLoadException">A version mismatch occurs</exception>
        /// <exception cref="ArgumentNullException">fullFileName is null or empty</exception>
        public static CdeclHandle LoadLibrary(string fullFileName, FileVersionInfo version = null)
        {
            if (String.IsNullOrWhiteSpace(fullFileName))
                throw new ArgumentNullException("fullFileName");
            if (!File.Exists(fullFileName))
                throw new FileNotFoundException("File is missing.", fullFileName);

            string fileName = Path.GetFileName(fullFileName);

            if (null != version)
            {
                FileVersionInfo fileVersion = FileVersionInfo.GetVersionInfo(fullFileName);
                if (version != fileVersion)
                {                
                    throw new FileLoadException(
                        String.Format("Unable to load library <{0}> because a version mismatch occurs." + fileName));
                }
            }

            IntPtr ptr = Interop.LoadLibrary(fullFileName);
            if (ptr == IntPtr.Zero)
                throw new Win32Exception(String.Format("Unable to load library <{0}>.", fileName));
            
            return new CdeclHandle(ptr, fileName);
        }

        /// <summary>
        /// Loads an unmanaged library from filesystem
        /// </summary>
        /// <param name="codebaseType">type to analyze directory/codebase from</param>
        /// <param name="fileName">name(incl. extension) without path of the library</param>
        /// <param name="version">optional file version to check</param>
        /// <returns>handle to library</returns>
        /// <exception cref="FileNotFoundException">File is missing</exception>
        /// <exception cref="Win32Exception">Unable to load library</exception>
        /// <exception cref="IOException">A version mismatch occurs</exception>
        /// <exception cref="ArgumentNullException">a non-optional argument is null or empty</exception>
        public static CdeclHandle LoadLibrary(Type codebaseType, string fileName, FileVersionInfo version = null)
        {
            if (null == codebaseType)
                throw new ArgumentNullException("codebaseType");
            if (String.IsNullOrWhiteSpace(fileName))
                throw new ArgumentNullException("fileName");

            string location = codebaseType.Assembly.Location;
            string folderPath = Path.GetDirectoryName(location);
            string fullFileName = Path.Combine(folderPath, fileName);
            return LoadLibrary(fullFileName, version);
        }

        /// <summary>
        /// Free the library/dispose the instance
        /// </summary>
        /// <exception cref="Win32Exception">Unable to free library</exception>
        public void Dispose()
        {
            if (Underlying != IntPtr.Zero)
            {
                if (!Interop.FreeLibrary(Underlying))
                    throw new Win32Exception(String.Format("Unable to free library <{0}>.", Name));
                Underlying = IntPtr.Zero;
            }
        }
    }
}