﻿using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196295.aspx </remarks>
	[SupportByVersion("Word", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum WdRevisionsBalloonPrintOrientation
	{
		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdBalloonPrintOrientationAuto = 0,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdBalloonPrintOrientationPreserve = 1,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdBalloonPrintOrientationForceLandscape = 2
	}
}