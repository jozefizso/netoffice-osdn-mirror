﻿using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.AccessApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("Access", 12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("2E70527E-92D1-43CC-A57B-ED48BCCC711D"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface DispCustomControlInReportEvents
	{
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class DispCustomControlInReportEvents_SinkHelper : SinkHelper, DispCustomControlInReportEvents
	{
		#region Static
		
		public static readonly string Id = "2E70527E-92D1-43CC-A57B-ED48BCCC711D";
		
		#endregion
		
		#region Ctor

		public DispCustomControlInReportEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region DispCustomControlInReportEvents
		
		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}