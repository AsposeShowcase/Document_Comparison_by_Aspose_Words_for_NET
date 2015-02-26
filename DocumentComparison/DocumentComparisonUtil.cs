using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace DocumentComparison
{
    /// <summary>
    /// Class to compare two MS Word documents
    /// </summary>
    internal class DocumentComparisonUtil
    {
        /// <summary>
        /// Compare the two documents and save the result as a Word document
        /// </summary>
        /// <param name="document1">First document</param>
        /// <param name="document2">Second document</param>
        /// <param name="comparisonDocument">Comparison document</param>
        internal void Compare(string document1, string document2, string comparisonDocument, ref int added, ref int deleted)
        {
            added = 0;
            deleted = 0;

            // Load both documents in Aspose.Words
            Document doc1 = new Document(document1);
            Document doc2 = new Document(document2);
            Document docComp = new Document(document1);
            DocumentBuilder builder = new DocumentBuilder(docComp);

            // Get sections of each document
            SectionCollection sectionList1 = doc1.Sections;
            SectionCollection sectionList2 = doc2.Sections;
            SectionCollection sectionListComp = docComp.Sections;

            // Go through all sections of first document
            for (int iSection = 0; iSection < sectionList1.Count; iSection++)
            {
                Section section1 = sectionList1[iSection];
                Section sectionComp = sectionListComp[iSection];
                // If second document does not have the same no of section, then no need to compare
                if (iSection >= sectionList2.Count)
                    break;
                // Get the paragraphs of each document
                ParagraphCollection paragraphs1 = section1.Body.Paragraphs;
                ParagraphCollection paragraphsComp = sectionComp.Body.Paragraphs;

                Section section2 = sectionList2[iSection];
                ParagraphCollection paragraphs2 = section2.Body.Paragraphs;

                // Loop through paragraphs for first document
                for (int iPara1 = 0; iPara1 < paragraphs1.Count; iPara1++)
                {
                    // Get the text from first document
                    string text1 = paragraphs1[iPara1].ToString(SaveFormat.Text);
                    string text2 = "";
                    Paragraph para2 = null;

                    // If first document has more paragraphs, let comparison text from second document be empty
                    if (iPara1 >= paragraphs2.Count)
                        text2 = ""; // para2 will be null
                    else
                    {
                        text2 = paragraphs2[iPara1].ToString(SaveFormat.Text);
                        // Get the reference of paragraph from second document
                        para2 = paragraphs2[iPara1];
                    }

                    List<Diff> differences = GetDiffList(text1, text2);

                    // Update the paragraph in the comparison document
                    Paragraph para = paragraphsComp[iPara1];

                    //Console.WriteLine("Para: " + para.ToString(SaveFormat.Text));
                    builder.MoveToParagraph(iPara1, 0);
                    UpdateParagraph(builder, para, para2, differences, ref added, ref deleted);
                }
            }

            docComp.Save(comparisonDocument);
        }

        /// <summary>
        /// Get the list of differences for two text strings
        /// </summary>
        /// <param name="text1"></param>
        /// <param name="text2"></param>
        /// <returns></returns>
        private List<Diff> GetDiffList(string text1, string text2)
        {
            diff_match_patch diffTest = new diff_match_patch();
            List<Diff> diffList = diffTest.diff_main(text1, text2);
            diffTest.diff_cleanupSemantic(diffList);
            Console.WriteLine("No. of differences: " + diffList.Count);
            foreach (Diff diff in diffList)
            {
                Console.WriteLine(diff.operation + ": " + diff.text);
            }
            return diffList;
        }

        /// <summary>
        /// Split run method, to divide one run into two runs
        /// </summary>
        /// <param name="run"></param>
        /// <param name="position"></param>
        /// <returns></returns>
        internal Run SplitRun(Run run, int position)
        {
            Run afterRun = (Run)run.Clone(true);
            afterRun.Text = run.Text.Substring(position);
            run.Text = run.Text.Substring(0, position);
            run.ParentNode.InsertAfter(afterRun, run);
            return afterRun;
        }

        /// <summary>
        /// Update paragraph and runs contained in it
        /// </summary>
        /// <param name="builder">DocumentBuilder instance</param>
        /// <param name="para">The paragraph which is to be updated</param>
        /// <param name="differences">List of differences</param>
        private void UpdateParagraph(DocumentBuilder builder, Paragraph para, Paragraph para2, List<Diff> differences, ref int added, ref int deleted)
        {
            int offset = 0;
            int insertOffset = 0; // For insertion only, incremented at insert and equal operations
            foreach (Diff diff in differences)
            {
                diff.text = diff.text.Replace("\r", "").Replace("\n", "");
                switch (diff.operation)
                {
                    case Operation.EQUAL:
                        // Find the run that is at position, equal to the length of text
                        //int runSplitOffset = 0;
                        //Run run = FindRun(para, diff.text.Length + offset, ref runSplitOffset);
                        offset += diff.text.Length;
                        insertOffset += diff.text.Length;
                        //Run nextRun = SplitRun(run, diff.text.Length);
                        //Run nextRun = SplitRun(run, runSplitOffset);
                        //Console.WriteLine("Run 1: " + run.Text);
                        //Console.WriteLine("Run 2: " + nextRun.Text);
                        //nextRun.Text = "<**>" + nextRun.Text;
                        //offset += 4;
                        break;
                    case Operation.INSERT:
                        added += 1;
                        // Insert the text at the current position + 1
                        int runSplitOffsetI = 0;
                        if (para.Runs.Count > 0)
                        {
                            Run runI = FindRun(para, offset + 1, ref runSplitOffsetI);
                            // Add a run before it, to add inserted text
                            // TODO
                            if (runSplitOffsetI - 1 > runI.Text.Length)
                                runSplitOffsetI = runI.Text.Length;
                            Run insertRun = SplitRun(runI, runSplitOffsetI - 1);
                            //runI.Text = diff.text;
                            //runI.Font.Color = Color.Blue;
                            //runI.Font.StrikeThrough = false;
                            //// Reset font of the split run
                            //insertRun.Font.ClearFormatting();
                            //Console.WriteLine("RunI 1: " + runI.Text);
                            //Console.WriteLine("RunI 2: " + insertRun.Text);

                            // Get the list of runs from the second document
                            // All these will be inserted in comparison document
                            List<Run> runsI = SearchParagraph(para2, insertOffset, diff.text.Length);
                            for (int iRun = runsI.Count - 1; iRun >= 0; iRun--)
                            {
                                Run runINew = runsI[iRun];
                                runINew.Font.Color = Color.Blue;
                                runINew.Font.StrikeThrough = false;

                                runI.ParentNode.InsertAfter(runI.Document.ImportNode(runINew, true), runI);
                            }

                            //runI.Remove();
                        }

                        offset += diff.text.Length;
                        insertOffset += diff.text.Length;

                        break;
                    case Operation.DELETE:
                        deleted += 1;
                        // Find the run that is at position, equal to the length of text
                        //int runSplitOffsetD = 0;
                        //Run runD = FindRun(para, diff.text.Length + offset, ref runSplitOffsetD);
                        List<Run> runsD = SearchParagraph(para, offset, diff.text.Length);
                        foreach (Run runD in runsD)
                        {
                            runD.Font.StrikeThrough = true;
                            runD.Font.Color = Color.Red;
                        }
                        offset += diff.text.Length;
                        //offset += runSplitOffsetD;
                        //Run nextRunD = SplitRun(runD, diff.text.Length);
                        //Run nextRunD = SplitRun(runD, runSplitOffsetD);
                        //Console.WriteLine("RunD 1: " + runD.Text);
                        //Console.WriteLine("RunD 2: " + nextRunD.Text);
                        break;
                }
            }
        }

        /// <summary>
        /// Find run that has the text of specified length
        /// </summary>
        /// <param name="para"></param>
        /// <param name="searchTextLength"></param>
        /// <returns></returns>
        private Run FindRun(Paragraph para, int searchTextLength, ref int runSplitOffset)
        {
            int runLength = 0;
            int runCount = 0;
            foreach (Run run in para.Runs)
            {
                runCount++;

                runLength += run.Text.Length;
                // If it is last run or the length is equal or greater
                if (runLength >= searchTextLength || runCount == para.Runs.Count)
                {
                    // If run is big and search text is small, offset = length of search text
                    if (run.Text.Length >= searchTextLength)
                        runSplitOffset = searchTextLength;
                    else //if (run.Text.Length < searchTextLength)
                        runSplitOffset = searchTextLength - runLength + run.Text.Length;
                    return run;
                }
            }
            return null;
        }

        /// <summary>
        /// Search a paragraph for text, works like a substring() method
        /// </summary>
        /// <param name="para"></param>
        /// <param name="startIndex"></param>
        /// <param name="length"></param>
        /// <returns>Returns a list of runs that will contain only the searched text</returns>
        public List<Run> SearchParagraph(Paragraph para, int startIndex, int length)
        {
            // Create an empty list of runs
            List<Run> runs = new List<Run>();

            // Split at start index
            SplitRunInsideParagraph(para, startIndex);
            // Again split at start index + length
            SplitRunInsideParagraph(para, startIndex + length);

            // After doing the necessary splitting, we now have perfect cuts (runs)
            // Now just return the list of runs
            int totalLength = 0;
            foreach (Run run in para.Runs)
            {
                if (totalLength >= startIndex &&
                    totalLength < startIndex + length)
                    runs.Add(run);

                totalLength += run.Text.Length;
            }

            return runs;
        }

        private void SplitRunInsideParagraph(Paragraph para, int position)
        {
            int totalLength = 0;
            int previousLength = 0;
            // Find the run that contains our start index
            foreach (Run run in para.Runs)
            {
                // Total Length of runs, till the previous run
                previousLength = totalLength;

                // Update current running total
                totalLength += run.Text.Length;
                if (totalLength >= position + 1)
                {
                    //Console.WriteLine("Found the start index.");
                    if (previousLength == position)
                    {
                        //Console.WriteLine("No need to split");
                        break; // Already cut at start index, no need to split
                    }
                    else
                    {
                        int splitPosition = position - previousLength;
                        //Console.WriteLine("Run split at " + splitPosition);
                        SplitRun(run, splitPosition);
                        break;
                    }
                }
            }
        }
    }
}