﻿using System;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.MSProjectApi.Tools.Utils
{
    /// <summary>
    /// Various helper for common tasks
    /// </summary>
    public class CommonUtils : NetOffice.OfficeApi.Tools.Contribution.CommonUtils
    {
        #region Fields

        private MSProjectApi.Application _ownerApplication;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the application
        /// </summary>
        /// <param name="application">owner application</param>
        public CommonUtils(MSProjectApi.Application application) : base(application)
        {
            _ownerApplication = application;
        }

        /// <summary>
        /// Creates an instance of the application
        /// </summary>
        /// <param name="application">owner application</param>
        /// <param name="ownerAssembly">owner assembly</param>
        public CommonUtils(MSProjectApi.Application application, Assembly ownerAssembly) : base(application, ownerAssembly)
        {
            if (null == application)
                throw new ArgumentNullException("application");

            _ownerApplication = application;
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">addin owner instance</param>
        /// <param name="isAutomation">host application is started for automation</param>
        internal CommonUtils(NetOffice.Tools.COMAddinBase owner, bool isAutomation) : base(owner, isAutomation)
        {

        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">addin owner</param>
        /// <param name="isAutomation">indicates the host application is currently in automation</param>
        /// <param name="ownerAssembly">owner application</param>
        internal CommonUtils(NetOffice.Tools.COMAddinBase owner, bool isAutomation, Assembly ownerAssembly) : base(owner, isAutomation, ownerAssembly)
        {

        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">addin owner</param>
        /// <param name="ownerType">type information from addin owner</param>
        /// <param name="isAutomation">indicates the host application is currently in automation</param>
        /// <param name="ownerAssembly">owner application</param>
        internal CommonUtils(NetOffice.Tools.COMAddinBase owner, Type ownerType, bool isAutomation, Assembly ownerAssembly) : base(owner, ownerType, isAutomation, ownerAssembly)
        {

        }

        #endregion

        #region Properties

        /// <summary>
        /// Encapsulate the owner application to make accessible for child utils
        /// </summary>
        public new MSProjectApi.Application OwnerApplication
        {
            get
            {
                return base.OwnerApplication as MSProjectApi.Application;
            }
        }

        #endregion
    }
}
