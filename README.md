CustomTaskPaneIssue
===================

##### Sample project for an issue with custom task panes in Word 2010.

The issue is that when opening an existing document, or when creating a new document, Word stops refreshing its built-in task panes (i.e. the Navigation pane, Styles pane, Apply Styles pane and Reveal Formatting pane) when an add-in adds its own custom task pane.

### Update (October 2015)

With the help of this repro, Microsoft has fixed the problem with **[KB3055034](https://support.microsoft.com/en-us/help/3055034/october-13-2015-update-for-office-2010-kb3055034)**:

> #### Improvements and fixes
> Fixes the following issue:
>   When a custom task pane is displayed in a document in Microsoft Word 2010, information isn't updated in built-in task panes. For example, the formatting of selected text isn't updated on the Reveal Formatting task pane.

## Steps to reproduce the behavior

1. Create a new VSTO Word 2010 solution using Visual Studio.
2. Add a user control UserControl1 (Windows Forms) to the project and add a button to the user control.
3. Add event handlers for the Application.NewDocument and Application.DocumentOpen events in the ThisAddin class
4. Add a custom task pane showing UserControl1 in the following event handlers: ThisAddin.Startup, Application.NewDocument and Application.DocumentOpen (see sample code below).
5. Set the DockPosition of the task pane to Office.MsoCTPDockPosition.msoCTPDockPositionLeft (the dock position actually makes a difference on how severe the bug is, see the further findings below)
6. Build the solution.
7. Start Word 2010; Word displays an empty new document with the task pane.
8. Open the Navigation pane, the Styles pane, the Apply Styles pane and the Reveal Formatting pane.
9. Select File -> New -> Blank Document (this creates a second window "Document2")
10. In the new document, type some text and apply style "Heading 1"
11. The Navigation pane, Styles pane, Apply Styles pane and Reveal Formatting pane are not updated as they normally would be without the add-in.
12. Click on the button in the custom task pane. Now the Navigation pane and Styles pane are refreshed. It seems that giving focus to the custom task pane triggers a refresh of the built-in task panes.

## Expected behavior

The display of the custom task pane should not have any influence on how Word refreshes its built-in task panes.

Further findings:

* If the custom task pane is docked to the left, the Navigation pane, Styles pane, Apply Styles pane and Reveal Formatting pane are not refreshed. However, if I change the DockPosition of the custom task pane to the right, the Navigation pane, Styles pane and Reveal Formatting pane are refreshed, but the Apply Styles pane still is not refreshed.
* Word 2013 does not exhibit this problem, everything works as expected.
