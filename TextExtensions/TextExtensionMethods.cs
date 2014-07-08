using EnvDTE;
using EnvDTE80;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace TextExtensions
{
    internal sealed class TextExtensionMethods
    {
        private static string GetClipboardText()
        {
            var clipboardData = Clipboard.GetDataObject();
            if (clipboardData == null) return string.Empty;
            return (string)clipboardData.GetData(DataFormats.StringFormat);
        }

        internal static bool HasActiveDocument(DTE2 applicationObject)
        {
            if (applicationObject == null) return false;
            if (applicationObject.ActiveDocument == null) return false;
            return !applicationObject.ActiveDocument.ReadOnly;
        }

        internal static bool CanPaste(DTE2 applicationObject)
        {
            if (applicationObject == null) return false;
            if (applicationObject.ActiveDocument == null) return false;
            if (applicationObject.ActiveDocument.ReadOnly) return false;
            return !string.IsNullOrEmpty(GetClipboardText());
        }

        internal static bool CanSort(DTE2 applicationObject)
        {
            if (applicationObject == null) return false;
            if (applicationObject.ActiveDocument == null) return false;
            if (applicationObject.ActiveDocument.ReadOnly) return false;
            var selection = (TextSelection)applicationObject.ActiveDocument.Selection;
            return selection.TextRanges.Count > 1;
        }

        internal static bool HasSelection(DTE2 applicationObject)
        {
            if (applicationObject == null) return false;
            if (applicationObject.ActiveDocument == null) return false;
            if (applicationObject.ActiveDocument.ReadOnly) return false;
            var selection = (TextSelection)applicationObject.ActiveDocument.Selection;
            return selection.TextRanges.Count >= 1;
        }

        // Method to append selected text with lines from the clipboard. This method performs the same
        // whether or not the selection region is rectangular.
        internal static void PasteAppend(DTE2 applicationObject)
        {
            // See if paste operation is feasible.
            if (!CanPaste(applicationObject)) return;

            // Start undo context immediately.
            applicationObject.UndoContext.Open("PasteAppend");

            try
            {
                // Grab selection region in current document or the line the cursor is on.
                var selection = (TextSelection)applicationObject.ActiveDocument.Selection;

                // Array of seperators to split the clipboard text.
                var seperators = new[] { "\r\n", "\r", "\n" };

                // Contains text from the clipboard.
                var clipboardText = GetClipboardText();

                // Contains line(s) from the clipboard as an array of strings.
                var clipboardTextLines = clipboardText.Split(seperators, StringSplitOptions.RemoveEmptyEntries);

                // Iterator for clipboardTextLines
                var i = clipboardTextLines.Length;

                foreach (TextRange range in selection.TextRanges)
                {
                    // Ensure we are working with the entire line.
                    range.StartPoint.StartOfLine();
                    range.EndPoint.EndOfLine();

                    // If the line actually contains text and is not empty and the last character
                    // found within the line is a whitespace character then move the end point to
                    // the character found after the last non-whitespace character on the line.
                    if (!string.IsNullOrEmpty(range.StartPoint.GetText(range.EndPoint).Trim()))
                    {
                        while (char.IsWhiteSpace(range.EndPoint.GetText(1), 0))
                        {
                            range.EndPoint.CharLeft();
                        }
                        range.EndPoint.CharRight();
                    }

                    // Insert the current clipboard line into the EndPoint.
                    range.EndPoint.Insert(clipboardTextLines[i++ % clipboardTextLines.Length]);
                }
            }
            finally
            {
                applicationObject.UndoContext.Close();
            }
        }

        // Method to prepend selected text with lines from the clipboard. This method performs the same
        // whether or not the selection region is rectangular.
        internal static void PastePrepend(DTE2 applicationObject)
        {
            // See if paste operation is feasible.
            if (!CanPaste(applicationObject)) return;

            // Start undo context immediately.
            applicationObject.UndoContext.Open("PastePrepend");

            try
            {
                // Grab selection region in current document or the line the cursor is on.
                var selection = (TextSelection)applicationObject.ActiveDocument.Selection;

                // Array of seperators to split the clipboard text.
                var seperators = new[] { "\r\n", "\r", "\n" };

                // Contains text from the clipboard.
                var clipboardText = GetClipboardText();

                // Contains line(s) from the clipboard as an array of strings.
                var clipboardTextLines = clipboardText.Split(seperators, StringSplitOptions.RemoveEmptyEntries);

                // Iterator for clipboardTextLines
                var i = clipboardTextLines.Length;

                foreach (TextRange range in selection.TextRanges)
                {
                    // Ensure we are working with the entire line.
                    range.StartPoint.StartOfLine();
                    range.EndPoint.EndOfLine();

                    // If the line actually contains text and is not empty and the first character
                    // found within the line is a whitespace character then move the start point to
                    // the first non-whitespace character on the line.
                    if (!string.IsNullOrEmpty(range.StartPoint.GetText(range.EndPoint).Trim()))
                        while (char.IsWhiteSpace(range.StartPoint.GetText(1), 0))
                            range.StartPoint.CharRight();

                    // Insert the current clipboard line into the StartPoint.
                    range.StartPoint.Insert(clipboardTextLines[i++ % clipboardTextLines.Length]);
                }
            }
            finally
            {
                applicationObject.UndoContext.Close();
            }
        }

        // Method to replace selected text with lines from the clipboard. If the selection
        // region is not rectangular, this method may give undesired results.
        internal static void PasteReplace(DTE2 applicationObject)
        {
            // See if paste operation is feasible.
            if (!CanPaste(applicationObject)) return;

            // Start undo context immediately.
            applicationObject.UndoContext.Open("PasteReplace");

            try
            {
                // Grab selection region in current document or the line the cursor is on.
                var selection = (TextSelection)applicationObject.ActiveDocument.Selection;

                // Array of seperators to split the clipboard text.
                var seperators = new[] { "\r\n", "\r", "\n" };

                // Contains text from the clipboard.
                var clipboardText = GetClipboardText();

                // Contains line(s) from the clipboard as an array of strings.
                var clipboardTextLines = clipboardText.Split(seperators, StringSplitOptions.RemoveEmptyEntries);

                // Iterator for clipboardTextLines
                var i = clipboardTextLines.Length;

                // Replace the current selection region with the current clipboard line.
                foreach (TextRange range in selection.TextRanges)
                    range.StartPoint.ReplaceText(range.EndPoint, clipboardTextLines[i++ % clipboardTextLines.Length],
                        (int)vsEPReplaceTextOptions.vsEPReplaceTextKeepMarkers);
            }
            finally
            {
                applicationObject.UndoContext.Close();
            }
        }

        // Method to sort lines of text from the selection region. This method performs the same
        // whether or not the selection region is rectangular.
        internal static void SortLines(DTE2 applicationObject)
        {
            // See if sort operation is feasible.
            if (!CanSort(applicationObject)) return;

            // Start undo context immediately.
            applicationObject.UndoContext.Open("SortLines");

            try
            {
                // Grab selection region in current document or the line the cursor is on.
                var selection = (TextSelection)applicationObject.ActiveDocument.Selection;

                // Collections to store list of text lines and ranges while sorting.
                var list = new List<string>();
                var rangeList = new List<TextRange>();
                var firstWordPositionFound = false;
                var spacingToFirstWord = string.Empty;

                foreach (TextRange range in selection.TextRanges)
                {
                    // Ensure we are working with the entire line.
                    range.StartPoint.StartOfLine();
                    range.EndPoint.EndOfLine();

                    // store trimmed line data
                    var str = range.StartPoint.GetText(range.EndPoint).Trim();

                    // If the line actually contains text and is not empty store the line
                    // data to sort along with a set of points to replace the sorted text.
                    if (!string.IsNullOrEmpty(str))
                    {
                        if (!firstWordPositionFound)
                        {
                            range.EndPoint.StartOfLine();
                            while (char.IsWhiteSpace(range.EndPoint.GetText(1), 0))
                                range.EndPoint.CharRight();

                            spacingToFirstWord = range.StartPoint.GetText(range.EndPoint);
                            // Reset the EndPoint back to the end of the line.
                            range.EndPoint.EndOfLine();
                            firstWordPositionFound = true;
                        }


                        list.Add(str);
                        rangeList.Add(range);
                    }
                    else
                    {
                        // Delete the line.
                        range.EndPoint.LineDown();
                        range.EndPoint.StartOfLine();
                        range.StartPoint.Delete(range.EndPoint);
                    }
                }

                // Perform default sorting on the data within the selction area.
                list.Sort();

                // Iterator for list.
                var i = 0;

                // Replace the current selection region with the current padded list line.
                foreach (var range in rangeList)
                    range.StartPoint.ReplaceText(range.EndPoint, spacingToFirstWord + list[i++],
                        (int)vsEPReplaceTextOptions.vsEPReplaceTextKeepMarkers);
            }
            finally
            {
                applicationObject.UndoContext.Close();
            }
        }

        // Method to sort lines of text within a rectangular selection region. If the selection
        // region is not rectangular, this method may give undesired results.
        internal static void SortSelection(DTE2 applicationObject)
        {
            // See if sort operation is feasible.
            if (!CanSort(applicationObject)) return;

            // Start undo context immediately.
            applicationObject.UndoContext.Open("SortSelection");

            try
            {
                // Grab selection region in current document or the line the cursor is on.
                var selection = (TextSelection)applicationObject.ActiveDocument.Selection;

                // Collections to store list of text lines and ranges while sorting.
                var list = new List<string>();
                var rangeList = new List<TextRange>();

                foreach (TextRange range in selection.TextRanges)
                {
                    // store trimmed line data
                    var str = range.StartPoint.GetText(range.EndPoint).Trim();

                    // If the line actually contains text and is not empty store the line
                    // data to sort along with a set of points to replace the sorted text.
                    if (!string.IsNullOrEmpty(str))
                    {
                        list.Add(str);
                        rangeList.Add(range);
                    }
                    else
                    {
                        // Delete the line.
                        range.EndPoint.LineDown();
                        range.EndPoint.StartOfLine();
                        range.StartPoint.Delete(range.EndPoint);
                    }
                }

                // Perform default sorting on the data within the selction area.
                list.Sort();

                // Iterator for list.
                var i = 0;

                // Replace the current selection region with the current list line.
                foreach (var range in rangeList)
                    range.StartPoint.ReplaceText(range.EndPoint, list[i++],
                        (int)vsEPReplaceTextOptions.vsEPReplaceTextKeepMarkers);
            }
            finally
            {
                applicationObject.UndoContext.Close();
            }
        }

        // Method to convert text within a selection region to a capitalized version of that text.
        internal static void SelectionCapitalize(DTE2 applicationObject)
        {
            // See if toUpper operation is feasible.
            if (!HasSelection(applicationObject)) return;

            // Start undo context immediately.
            applicationObject.UndoContext.Open("SelectionCapitalize");

            try
            {
                // Grab selection region in current document or the line the cursor is on.
                var selection = (TextSelection)applicationObject.ActiveDocument.Selection;

                // Replace the current selection region with capitalized version.
                foreach (TextRange range in selection.TextRanges)
                    range.StartPoint.ChangeCase(range.EndPoint, vsCaseOptions.vsCaseOptionsCapitalize);
            }
            finally
            {
                applicationObject.UndoContext.Close();
            }
        }

        // Method to convert text within a selection region to a lower-case version of that text.
        internal static void SelectionToLower(DTE2 applicationObject)
        {
            // See if toUpper operation is feasible.
            if (!HasSelection(applicationObject)) return;

            // Start undo context immediately.
            applicationObject.UndoContext.Open("SelectionToLower");

            try
            {
                // Grab selection region in current document or the line the cursor is on.
                var selection = (TextSelection)applicationObject.ActiveDocument.Selection;
                selection.ChangeCase(vsCaseOptions.vsCaseOptionsLowercase);
            }
            finally
            {
                applicationObject.UndoContext.Close();
            }
        }

        // Method to convert text within a selection region to an upper-case version of that text.
        internal static void SelectionToUpper(DTE2 applicationObject)
        {
            // See if toUpper operation is feasible.
            if (!HasSelection(applicationObject)) return;

            // Start undo context immediately.
            applicationObject.UndoContext.Open("SelectionToUpper");

            try
            {
                // Grab selection region in current document or the line the cursor is on.
                var selection = (TextSelection)applicationObject.ActiveDocument.Selection;
                selection.ChangeCase(vsCaseOptions.vsCaseOptionsUppercase);
            }
            finally
            {
                applicationObject.UndoContext.Close();
            }
        }
    }
}
