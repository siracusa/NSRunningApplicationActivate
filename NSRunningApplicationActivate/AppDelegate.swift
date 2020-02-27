//
//  AppDelegate.swift
//  NSRunningApplicationActivate
//
//  Created by John Siracusa on 2/27/20.
//  Copyright © 2020 John Siracusa. All rights reserved.
//

import Cocoa

@NSApplicationMain
class AppDelegate: NSObject, NSApplicationDelegate {

    @IBOutlet weak var window: NSWindow!


    func applicationDidFinishLaunching(_ aNotification: Notification) {
        // Insert code here to initialize your application
    }

    func applicationWillTerminate(_ aNotification: Notification) {
        // Insert code here to tear down your application
    }

    @IBAction func activateOutlook(_ sender: Any) {
        let apps = NSWorkspace.shared.runningApplications.filter { $0.bundleIdentifier == "com.microsoft.Outlook" }
        
        if apps.count == 0 {
            let alert = NSAlert()
            alert.messageText = "Microsoft Outlook is not running."
            alert.informativeText = "Please launch Microsoft Outlook and try the button again."
            alert.alertStyle = .critical
            alert.runModal()
        }
        else if apps.count > 1 {
            let alert = NSAlert()
            alert.messageText = "Found more than one instance of Microsoft Outlook."
            alert.informativeText = "Using the first one…"
            alert.alertStyle = .informational
            alert.runModal()
        }
        else {
            if let app = apps.first {
                //
                // BEGIN IMPORTANT PART
                //

                app.activate(options: [.activateAllWindows, .activateIgnoringOtherApps])

                //
                // END IMPORTANT PART
                //
            }
            else {
                let alert = NSAlert()
                alert.messageText = "Could not get NSRunningApplication instance for Microsoft Outlook"
                alert.informativeText = "Bleh…"
                alert.alertStyle = .critical
                alert.runModal()
            }
        }
    }
}

