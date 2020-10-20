# Context

The [Office JavaScript API](https://docs.microsoft.com/en-us/office/dev/add-ins/reference/javascript-api-for-office) defines the [`saveAsync` method](https://docs.microsoft.com/en-us/javascript/api/office/office.settings?view=word-js-preview#saveasync-callback-) to persist settings into the document (e.g. the PowerPoint presentation).

This is an asynchronous method which triggers a callback once the setting has been persisted. In some cases, the callback is never called. This repo shows such a case where the callback is never called when a selector is present in the HTML.

# System information

This issue was reproduced with

- Microsoft PowerPoint for Office 365 MSO (16.0.12527.21096) 64-bit with a developer license
- Windows 10 Pro, version 2004, OS build 19041.572

# How to run this add-in

- Follow the instructions to [sideload an Office Add-in for testing](https://docs.microsoft.com/en-us/javascript/api/office/office.settings?view=word-js-preview#saveasync-callback-). This was tested using the "[Sideload Office Add-ins for testing from a network share](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)" guide. The manifest is located in this repo `./manifest.xml`
- To run the add-in, we used VSCode and the [Live Server extension](https://marketplace.visualstudio.com/items?itemName=ritwickdey.LiveServer).

If you encounter the issue "We can't open this add-in from localhost" when loading an Office Add-in or using Fiddler, you might have to run the following command in the command prompt as administrator ([source](https://docs.microsoft.com/en-us/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost))

```cmd
CheckNetIsolation LoopbackExempt -a -n="microsoft.win32webviewhost_cw5n1h2txyewy"
```

# How to reproduce the issue

Once the Live Server is running, do the following:

1. Open PowerPoint
2. Create a new presentation
3. Insert the "[dev] Counter" add-in
4. Once loaded, click on "Increment". The state should update and the logs should display "..."
5. Click on the `<select />` element and change its value.
6. Click on "Increment"

  - Expected: the `saveAsync` callback is called (as displayed in the logs)
  - Actual: the `saveAsync` callback is never called (it doesn't show up in logs).

You can also find a [video here](resources/2020-10-20_10-41-51.mp4) which shows the issue.

Notes:
- The `<select />` does not server a purpose other than reproducing this bug. It's not used in the code.
- Even though the `saveAsync` callback is not called, the state is still persisted. This can be verified by saving the presentation and opening it once again.
