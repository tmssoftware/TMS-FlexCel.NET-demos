// WARNING
//
// This file has been generated automatically by Xamarin Studio to store outlets and
// actions made in the UI designer. If it is removed, they will be lost.
// Manual changes to this file may not be handled correctly.
//
using Foundation;
using System.CodeDom.Compiler;

namespace FlexView
{
	[Register ("FlexViewViewController")]
	partial class FlexViewViewController
	{

		[Outlet]
		UIKit.UIBarButtonItem RandomizeButton { get; set; }

		[Outlet]
		UIKit.UIBarButtonItem ShareButton { get; set; }

		[Outlet]
		UIKit.UIWebView Viewer { get; set; }

		[Action ("RandomizeClick:")]
		[GeneratedCodeAttribute ("iOS Designer", "1.0")]
		partial void RandomizeClick (UIKit.UIBarButtonItem sender);

		[Action ("ShareClick:")]
		partial void ShareClick (UIKit.UIBarButtonItem sender);
		
		void ReleaseDesignerOutlets ()
		{
			if (RandomizeButton != null) {
				RandomizeButton.Dispose ();
				RandomizeButton = null;
			}

			if (ShareButton != null) {
				ShareButton.Dispose ();
				ShareButton = null;
			}

			if (Viewer != null) {
				Viewer.Dispose ();
				Viewer = null;
			}
		}
	}
}
