// WARNING
//
// This file has been generated automatically by Xamarin Studio to store outlets and
// actions made in the UI designer. If it is removed, they will be lost.
// Manual changes to this file may not be handled correctly.
//
using Foundation;
using System.CodeDom.Compiler;

namespace LangWars
{
	[Register ("LangWarsViewController")]
	partial class LangWarsViewController
	{
		[Outlet]
		UIKit.UISegmentedControl OfflineSwitch { get; set; }

		[Outlet]
		UIKit.UIActivityIndicatorView ProgressIndicator { get; set; }

		[Outlet]
		UIKit.UIWebView ResultsWindow { get; set; }

		[Outlet]
		UIKit.UIBarButtonItem ShareButton { get; set; }

		[Action ("FightClick:")]
		partial void FightClick (Foundation.NSObject sender);

		[Action ("ShareClick:")]
		partial void ShareClick (Foundation.NSObject sender);
		
		void ReleaseDesignerOutlets ()
		{
			if (OfflineSwitch != null) {
				OfflineSwitch.Dispose ();
				OfflineSwitch = null;
			}

			if (ResultsWindow != null) {
				ResultsWindow.Dispose ();
				ResultsWindow = null;
			}

			if (ProgressIndicator != null) {
				ProgressIndicator.Dispose ();
				ProgressIndicator = null;
			}

			if (ShareButton != null) {
				ShareButton.Dispose ();
				ShareButton = null;
			}
		}
	}
}
