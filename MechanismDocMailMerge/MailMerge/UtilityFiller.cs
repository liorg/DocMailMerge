using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.IO;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Validation;
 
namespace Guardian.Documents.MailMerge
{
    /*       *** thank's god ***         */
    /*
     *********************** current ver 1.0.0.8.3  ******************************
     */
    //version 1.0.0.8.3
    //1. fixed after
    // version 1.0.0.8.2
    // 1. header fixed problem on table
    // version 1.0.0.8.1
    //1. when no has data in all tablerow remove row
    // version 1.0.0.8
    //1. on function RemoveParentParagrph , fill to array tableCells any pargraph that's remove and has parent tableCell
    //2. on function FillWordFieldsInElement ,remove tablecell no has any paragraph
    // version 1.0.0.7
    // 1. when has paragraph contain text do not remove him (on body) RemoveParentParagrph
    // 2. remove rtl=false on  GetRunElementForTextRtlHandle
    // version 1.0.0.6
    // 1. default heb (on body)
    // version 1.0.0.5
    // 1. fixed parent is null
    // 2. font hebrew
    // version 1.0.0.4
    // 1. Remove Paragraph from mail merge field for prevent line to be empty  l.g -RemoveParentParagrph
    //version 1.0.0.3
    //1.handle datetime number sticky to month l.g - ConvertFieldCodes
    //2. decorate code
    // version 1.0.0.2:
    // 1. when handle rtl ltr on some context then create for each word run object and set rtl=false (disable) on english charcters -GetRunElementForTextRtlHandle
    // version 1.0.0.1:
    // 1. remove rtl from ConvertFieldCodes (fixed in GetRunElementForText)
    // 2. add handle on pre text mailmerge field (was error in opening word after fill merge)
    // 3 handle quetes when has on merger mail field exmple { MERGEFIELD  "MitlonenFullName" }
    /// <summary>
    /// Helper class for filling in data forms based on Word
    /// http://www.codeproject.com/Articles/38575/Fill-Mergefields-in-docx-Documents-without-Microso
    /// </summary>
    public class UtilityFiller
    {
        /// <summary>
        /// Regex used to parse MERGEFIELDs in the provided document.
        /// </summary>
        //        private static readonly Regex instructionRegEx =
        //            new Regex(
        //                        @"^[\s]*MERGEFIELD[\s]+(?<name>[#\w]*){1}               # This retrieves the field's name (Named Capture Group -> name)
        //                            [\s]*(\\\*[\s]+(?<Format>[\w]*){1})?                # Retrieves field's format flag (Named Capture Group -> Format)
        //                            [\s]*(\\b[\s]+[""]?(?<PreText>[^\\]*){1})?         # Retrieves text to display before field data (Named Capture Group -> PreText)
        //                                                                                # Retrieves text to display after field data (Named Capture Group -> PostText)
        //                            [\s]*(\\f[\s]+[""]?(?<PostText>[^\\]*){1})?",
        //                        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.ExplicitCapture | RegexOptions.IgnoreCase | RegexOptions.IgnorePatternWhitespace | RegexOptions.Singleline);
 
        // http://stackoverflow.com/questions/12206667/i-need-to-modify-a-word-mergefield-regular-expression
        // lior gr fixed because the old not support {MERGEFIELD  "xx"}  only { MERGEFIELD  xx} on field  exmple { MERGEFIELD  "MitlonenFullName" }
        private static readonly Regex instructionRegEx = new Regex(
                                   @"^[\s]*MERGEFIELD[\s]+""?(?<name>[#\w]*){1}""?           # This retrieves the field's name (Named Capture Group -> name) fixed l.g
                                              [\s]*(\\\*[\s]+(?<Format>[\w]*){1})?                # Retrieves field's format flag (Named Capture Group -> Format)
                                              [\s]*(\\b[\s]+[""]?(?<PreText>[^\\]*){1})?         # Retrieves text to display before field data (Named Capture Group -> PreText)
                                                                                        # Retrieves text to display after field data (Named Capture Group -> PostText)
                                              [\s]*(\\f[\s]+[""]?(?<PostText>[^\\]*){1})?",
                            RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.ExplicitCapture | RegexOptions.IgnoreCase | RegexOptions.IgnorePatternWhitespace | RegexOptions.Singleline);
 
        private static readonly string splitWhiteSpaceIncludeRegFormater = @"(?<=[\s+])";
 
 
        private static readonly string isEngCharcter = @"[A-Za-z]+";
 
        /// <summary>
        /// Change code from code project that's only have array of fields(lior G)
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="docx"></param>
        /// <param name="values"></param>
        internal static void GetWordReportPart(MemoryStream stream, WordprocessingDocument docx, Dictionary<string, string> values, string defaultHebDocBody = "David")
        {
 
 
            //  2010/08/01: addition
            ConvertFieldCodes(docx.MainDocumentPart.Document);
 
            // next : process all remaining fields in the main document //1.0.0.6
            FillWordFieldsInElement(values, docx.MainDocumentPart.Document, defaultHebDocBody);
 
            docx.MainDocumentPart.Document.Save();  // save main document back in package
 
            //// process header(s)
            foreach (HeaderPart hpart in docx.MainDocumentPart.HeaderParts)
            {
                //  2010/08/01: addition
                ConvertFieldCodes(hpart.Header);
                FillWordFieldsInElement(values, hpart.Header);
                hpart.Header.Save();    // save header back in package
            }
            // process footer(s)
            foreach (FooterPart fpart in docx.MainDocumentPart.FooterParts)
           {
                //  2010/08/01: addition
                ConvertFieldCodes(fpart.Footer);
                FillWordFieldsInElement(values, fpart.Footer);
                fpart.Footer.Save();    // save footer back in package
            }
        }
 
        /// <summary>
        /// Fills all the <see cref="SimpleFields"/> that are found in a given <see cref="OpenXmlElement"/>.
        /// </summary>
        /// <param name="values">The values to insert; keys should match the placeholder names, values are the data to insert.</param>
        /// <param name="element">The document element taht will contain the new values.</param>
        static void FillWordFieldsInElement(Dictionary<string, string> values, OpenXmlElement element, string defaultHebDoc = "")
        {
            string[] switches;
            string[] options;
            string[] formattedText;
 
            Dictionary<SimpleField, string[]> emptyfields = new Dictionary<SimpleField, string[]>();
 
            // First pass: fill in data, but do not delete empty fields.  Deletions silently break the loop.
            var list = element.Descendants<SimpleField>().ToArray();
            List<Paragraph> paragrphsChanged = new List<Paragraph>();
            foreach (var field in list)
            {
 
                var parent = field.Parent;
 
                string fieldname = GetFieldNameWithOptions(field, out switches, out options);
                if (!string.IsNullOrEmpty(fieldname))
                {
                    if (values.ContainsKey(fieldname)
                        && !string.IsNullOrEmpty(values[fieldname]))
                    {
                        //TODO: Fixed Before Merge Field
                        formattedText = ApplyFormatting(options[0], values[fieldname], options[1], options[2]);
 
                        // Prepend any text specified to appear before the data in the MergeField
                        if (!string.IsNullOrEmpty(options[1]))
                        {
 
                            // not working l.g
                            // field.Parent.InsertBeforeSelf<Paragraph>(GetPreOrPostParagraphToInsert(formattedText[2], field));
                            // start:replace with that
                            var text = formattedText[1];
                            Run runToInsert = GetRunElementForTextRtlHandle(text, field, defaultHebDoc); // GetRunElementForText(text, field, true);
                            var runprop = runToInsert.GetFirstChild<RunProperties>();
 
                            if (runprop != null && runprop.RightToLeftText != null)
                                runprop.RightToLeftText.Val = false;
 
                            parent.InsertBeforeSelf<Run>(runToInsert); //field.Parent.InsertBeforeSelf<Run>(runToInsert);
                            // end: replace l.g
                        }
                        // Append any text specified to appear after the data in the MergeField 
                        //start: 1.0.0.8.3
                        if (!string.IsNullOrEmpty(options[2]))
                        {
                            var text = formattedText[2];
                            Run runToInsert = GetRunElementForTextRtlHandle(text, field, defaultHebDoc);
                            var runprop = runToInsert.GetFirstChild<RunProperties>();
 
                            if (runprop != null && runprop.RightToLeftText != null)
                                runprop.RightToLeftText.Val = false;
 
                            parent.InsertAfterSelf<Run>(runToInsert);
                            //remove code  1.0.0.8.3
                            //// parent.InsertAfterSelf<Paragraph>(GetPreOrPostParagraphToInsert(formattedText[2], field));// field.Parent.InsertAfterSelf<Paragraph>(GetPreOrPostParagraphToInsert(formattedText[2], field));
                        }
                        //end: 1.0.0.8.3
                        // replace mergefield with text
                        parent.ReplaceChild<SimpleField>(GetRunElementForTextRtlHandle(formattedText[0], field, defaultHebDoc), field);// field.Parent.ReplaceChild<SimpleField>(GetRunElementForTextRtlHandle(formattedText[0], field), field);
 
                        var paragraphItemChange = GetFirstParent<Paragraph>(parent);
                        if (paragraphItemChange != null && !paragrphsChanged.Contains(paragraphItemChange))
                        {
                            paragrphsChanged.Add(paragraphItemChange);
                        }
 
                    }
                    else
                    {
                        // keep track of unknown or empty fields
                        emptyfields[field] = switches;
                    }
                }
            }
            //fix 1.0.0.8
            List<TableCell> tableCellRemoveParagraphs = new List<TableCell>();
            // second pass : clear empty fields
            foreach (KeyValuePair<SimpleField, string[]> kvp in emptyfields)
            {
                // if field is unknown or empty: execute switches and remove it from document !
                ExecuteSwitches(kvp.Key, kvp.Value);
                // 1.0.0.5 l.g
                RemoveParentParagrph(kvp.Key, paragrphsChanged, tableCellRemoveParagraphs);
 
                kvp.Key.Remove();
 
            }
            //fixed 1.0.0.8
            DelCellHasNoPargraphs(tableCellRemoveParagraphs);
        }
 
        /// <summary>
        /// Since MS Word 2010 the SimpleField element is not longer used. It has been replaced by a combination of
        /// Run elements and a FieldCode element. This method will convert the new format to the old SimpleField-compliant
        /// format.
        /// </summary>
        /// <param name="mainElement"></param>
        static void ConvertFieldCodes(OpenXmlElement mainElement)
        {
            //  search for all the Run elements
            Run[] runs = mainElement.Descendants<Run>().ToArray();
            if (runs.Length == 0) return;
            string[] allDateFormats = DateTimeFormatInfo.CurrentInfo.GetAllDateTimePatterns('d');
            List<string> datesFormat = new List<string>();
            // Print out july28 in all DateTime formats using the default culture.
            foreach (string format in allDateFormats)
            {
                datesFormat.Add(format);
            }
 
            var newfields = new Dictionary<Run, Run[]>();
 
            int cursor = 0;
            do
            {
                Run run = runs[cursor];
 
                if (run.HasChildren && run.Descendants<FieldChar>().Count() > 0
                    && (run.Descendants<FieldChar>().First().FieldCharType & FieldCharValues.Begin) == FieldCharValues.Begin)
                {
                    List<Run> innerRuns = new List<Run>();
                    innerRuns.Add(run);
 
                    //  loop until we find the 'end' FieldChar
                    bool found = false;
                    string instruction = null;
                    RunProperties runprop = null;
                    do
                    {
                        cursor++;
                        run = runs[cursor];
 
                        innerRuns.Add(run);
                        if (run.HasChildren && run.Descendants<FieldCode>().Count() > 0)
                            instruction += run.GetFirstChild<FieldCode>().Text;
                        if (run.HasChildren && run.Descendants<FieldChar>().Count() > 0
                            && (run.Descendants<FieldChar>().First().FieldCharType & FieldCharValues.End) == FieldCharValues.End)
                        {
                            found = true;
                        }
                        if (run.HasChildren && run.Descendants<RunProperties>().Count() > 0)
                            runprop = run.GetFirstChild<RunProperties>();
                        // Fixed by lior G
                        if (runprop != null && runprop.RightToLeftText != null)
                        {
                            //  runprop.RightToLeftText.Val = false;
 
                        }
                    } while (found == false && cursor < runs.Length);
 
                    //  something went wrong : found Begin but no End. Throw exception
                    if (!found)
                        throw new Exception("Found a Begin FieldChar but no End !");
 
                    if (!string.IsNullOrEmpty(instruction))
                    {
 
                        Run newrun = new Run();
                        // Jewish Calander Handle when is dynamic update
                        // lior grossman
                        // if (instruction.Contains("DATE \\@ \"dd/MM/yyyy\" \\h "))
                        bool isFoundFormatDate = false;
                        foreach (var format in datesFormat)
                        {
                            if (instruction.Contains(format))
                            {
                                isFoundFormatDate = true;
                                break;
                            }
                        }
                        if (isFoundFormatDate)
                        {
                            // Fixed jewish calander --lior
                            //  build new Run containing a SimpleField
                            Paragraph paragraphDate = null;
                            if (runprop != null && runprop.Parent != null && runprop.Parent.Parent != null)
                                paragraphDate = runprop.Parent.Parent as Paragraph;
 
                            else if (innerRuns.Any() && (innerRuns[0] != null && innerRuns[0].Parent != null))
                                paragraphDate = runprop.Parent.Parent as Paragraph;
 
                            if (paragraphDate != null)
                            {
                                var texts = paragraphDate.Descendants<Text>();
                                // start: handle datetime number sticky to month l.g
                                foreach (var text in texts)
                                {
                                    if (text == null)
                                        continue;
 
                                    Run rWord = new Run();
                                    var txtTemp = new Text(text.InnerText);
                                    txtTemp.Space = SpaceProcessingModeValues.Preserve;
                                    if (runprop != null)
                                    {
                                        rWord.AppendChild(runprop.CloneNode(true));
                                    }
                                    rWord.AppendChild<Text>(txtTemp);
                                    newrun.AppendChild<Run>(rWord);
                                }
                                // end: handle datetime number sticky to month l.g
                                // must be after text set all properties
                                if (runprop != null)
                                {
                                    var propCalder = runprop.CloneNode(true);
                                    newrun.AppendChild(propCalder);
                                }
                            }
                            else
                            {
                                // must be befor simplefield set all properties
                                if (runprop != null)
                                    newrun.AppendChild(runprop.CloneNode(true));
                                SimpleField simplefield = new SimpleField();
                                simplefield.Instruction = instruction;
                                newrun.AppendChild(simplefield);
                            }
 
                        } //end is field date fromat
                        else
                        {
                            // must be befor simplefield set all properties
                            if (runprop != null)
                                newrun.AppendChild(runprop.CloneNode(true));
                            // no field date format
                            SimpleField simplefield = new SimpleField();
                            simplefield.Instruction = instruction;
                            newrun.AppendChild(simplefield);
                        }
                        newfields.Add(newrun, innerRuns.ToArray());
                    }
                }
                cursor++;
            } while (cursor < runs.Length);
 
            //  replace all FieldCodes by old-style SimpleFields
            foreach (KeyValuePair<Run, Run[]> kvp in newfields)
            {
                kvp.Value[0].Parent.ReplaceChild(kvp.Key, kvp.Value[0]);
                for (int i = 1; i < kvp.Value.Length; i++)
                    kvp.Value[i].Remove();
            }
        }
 
        [Obsolete("replace with new GetRunElementForText ", false)]
        static Run GetRunElementForTextOld(string text, SimpleField placeHolder)
        {
            string rpr = null;
            if (placeHolder != null)
            {
                foreach (RunProperties placeholderrpr in placeHolder.Descendants<RunProperties>())
                {
                    rpr = placeholderrpr.OuterXml;
                    break;  // break at first
                }
            }
 
            Run r = new Run();
            if (!string.IsNullOrEmpty(rpr))
            {
                r.Append(new RunProperties(rpr));
            }
 
            if (!string.IsNullOrEmpty(text))
            {
                // first process line breaks
                string[] split = text.Split(new string[] { "\n" }, StringSplitOptions.None);
                bool first = true;
                foreach (string s in split)
                {
                    if (!first)
                    {
                        r.Append(new Break());
                    }
 
                    first = false;
 
                    // then process tabs
                    bool firsttab = true;
                    string[] tabsplit = s.Split(new string[] { "\t" }, StringSplitOptions.None);
                    foreach (string tabtext in tabsplit)
                    {
                        if (!firsttab)
                        {
                            r.Append(new TabChar());
                        }
 
                        r.Append(new Text(tabtext));
                        firsttab = false;
                    }
                }
            }
 
            return r;
        }
 
        /// <summary>
        /// Returns a <see cref="Run"/>-openxml element for the given text.
        /// Specific about this run-element is that it can describe multiple-line and tabbed-text.
        /// The <see cref="SimpleField"/> placeholder can be provided too, to allow duplicating the formatting.
        /// </summary>
        /// <param name="text">The text to be inserted.</param>
        /// <param name="placeHolder">The placeholder where the text will be inserted.</param>
        /// <returns>A new <see cref="Run"/>-openxml element containing the specified text.</returns>
        internal static Run GetRunElementForText(string text, SimpleField placeHolder, bool preserveWhiteSpace = false)
        {
            string rpr = null;
            if (placeHolder != null)
            {
                var xdoc = XDocument.Parse((placeHolder.Parent).OuterXml.Replace(placeHolder.OuterXml, string.Empty));
                if (xdoc.Root != null)
                {
                    var xrpr = xdoc.Root.Elements().FirstOrDefault(x => x.Name.LocalName == "rPr");
 
                    if (xrpr != null)
                        rpr = xrpr.ToString();
                }
            }
 
            var r = new Run();
 
            if (!string.IsNullOrEmpty(rpr))
            {
                r.AppendChild(new RunProperties(rpr));
            }
 
            if (!string.IsNullOrEmpty(text))
            {
                var rprLtr = r.GetFirstChild<RunProperties>();
                // l.g fixed ltr in latin language
                bool isHasEngChar = Regex.IsMatch(text, isEngCharcter);//, @"[A-Za-z]+");  //l.g
                if (isHasEngChar && r.GetFirstChild<RunProperties>() != null)
                {
 
                    if (rprLtr.RightToLeftText == null)
                        rprLtr.RightToLeftText = new RightToLeftText();
                    rprLtr.RightToLeftText.Val = false;
                }
 
                //   var rRunProperties = r.GetFirstChild<RunProperties>();
                string[] split = text.Split(new string[] { "\n" }, StringSplitOptions.None);
                bool first = true;
                foreach (string s in split)
                {
                    if (!first)
                    {
                        r.Append(new Break());
                    }
 
                    first = false;
 
                    // then process tabs
                    bool firsttab = true;
                    string[] tabsplit = s.Split(new[] { "\t" }, StringSplitOptions.None);
                    foreach (string tabtext in tabsplit)
                    {
                        if (!firsttab)
                        {
                            r.Append(new TabChar());
                        }
                        var txt = new Text(tabtext);
                        txt.Space = SpaceProcessingModeValues.Preserve;
                        r.AppendChild<Text>(txt);
                        firsttab = false;
                    }
                }
            }
 
            return r;
        }
 
 
        /// <summary>
        /// Applies any formatting specified to the pre and post text as
        /// well as to fieldValue.
        /// </summary>
        /// <param name="format">The format flag to apply.</param>
        /// <param name="fieldValue">The data value being inserted.</param>
        /// <param name="preText">The text to appear before fieldValue, if any.</param>
        /// <param name="postText">The text to appear after fieldValue, if any.</param>
        /// <returns>The formatted text; [0] = fieldValue, [1] = preText, [2] = postText.</returns>
        /// <exception cref="">Throw if fieldValue, preText, or postText are null.</exception>
        static string[] ApplyFormatting(string format, string fieldValue, string preText, string postText)
        {
            string[] valuesToReturn = new string[3];
 
            if ("UPPER".Equals(format))
            {
                // Convert everything to uppercase.
                valuesToReturn[0] = fieldValue.ToUpper(CultureInfo.CurrentCulture);
                valuesToReturn[1] = preText.ToUpper(CultureInfo.CurrentCulture);
                valuesToReturn[2] = postText.ToUpper(CultureInfo.CurrentCulture);
            }
            else if ("LOWER".Equals(format))
            {
                // Convert everything to lowercase.
                valuesToReturn[0] = fieldValue.ToLower(CultureInfo.CurrentCulture);
                valuesToReturn[1] = preText.ToLower(CultureInfo.CurrentCulture);
                valuesToReturn[2] = postText.ToLower(CultureInfo.CurrentCulture);
            }
            else if ("FirstCap".Equals(format))
            {
                // Capitalize the first letter, everything else is lowercase.
                if (!string.IsNullOrEmpty(fieldValue))
                {
                    valuesToReturn[0] = fieldValue.Substring(0, 1).ToUpper(CultureInfo.CurrentCulture);
                    if (fieldValue.Length > 1)
                    {
                        valuesToReturn[0] = valuesToReturn[0] + fieldValue.Substring(1).ToLower(CultureInfo.CurrentCulture);
                    }
                }
 
                if (!string.IsNullOrEmpty(preText))
                {
                    valuesToReturn[1] = preText.Substring(0, 1).ToUpper(CultureInfo.CurrentCulture);
                    if (fieldValue.Length > 1)
                    {
                        valuesToReturn[1] = valuesToReturn[1] + preText.Substring(1).ToLower(CultureInfo.CurrentCulture);
                    }
                }
 
                if (!string.IsNullOrEmpty(postText))
                {
                    valuesToReturn[2] = postText.Substring(0, 1).ToUpper(CultureInfo.CurrentCulture);
                    if (fieldValue.Length > 1)
                    {
                        valuesToReturn[2] = valuesToReturn[2] + postText.Substring(1).ToLower(CultureInfo.CurrentCulture);
                    }
                }
            }
            else if ("Caps".Equals(format))
            {
                // Title casing: the first letter of every word should be capitalized.
                valuesToReturn[0] = ToTitleCase(fieldValue);
                valuesToReturn[1] = ToTitleCase(preText);
                valuesToReturn[2] = ToTitleCase(postText);
            }
            else
            {
                valuesToReturn[0] = fieldValue;
                valuesToReturn[1] = preText;
                valuesToReturn[2] = postText;
            }
 
            return valuesToReturn;
        }
 
        /// <summary>
        /// Title-cases a string, capitalizing the first letter of every word.
        /// </summary>
        /// <param name="toConvert">The string to convert.</param>
        /// <returns>The string after title-casing.</returns>
        static string ToTitleCase(string toConvert)
        {
            return ToTitleCaseHelper(toConvert, string.Empty);
        }
 
        /// <summary>
        /// Title-cases a string, capitalizing the first letter of every word.
        /// </summary>
        /// <param name="toConvert">The string to convert.</param>
        /// <param name="alreadyConverted">The part of the string already converted.  Seed with an empty string.</param>
        /// <returns>The string after title-casing.</returns>
        static string ToTitleCaseHelper(string toConvert, string alreadyConverted)
        {
            /*
             * Tail-recursive title-casing implementation.
             * Edge case: toConvert is empty, null, or just white space.  If so, return alreadyConverted.
             * Else: Capitalize the first letter of the first word in toConvert, append that to alreadyConverted and recur.
             */
            if (string.IsNullOrEmpty(toConvert))
            {
                return alreadyConverted;
            }
            else
            {
                int indexOfFirstSpace = toConvert.IndexOf(' ');
                string firstWord, restOfString;
 
                // Check to see if we're on the last word or if there are more.
                if (indexOfFirstSpace != -1)
                {
                    firstWord = toConvert.Substring(0, indexOfFirstSpace);
                    restOfString = toConvert.Substring(indexOfFirstSpace).Trim();
                }
                else
                {
                    firstWord = toConvert.Substring(0);
                    restOfString = string.Empty;
                }
 
                System.Text.StringBuilder sb = new StringBuilder();
 
                sb.Append(alreadyConverted);
                sb.Append(" ");
                sb.Append(firstWord.Substring(0, 1).ToUpper(CultureInfo.CurrentCulture));
 
                if (firstWord.Length > 1)
                {
                    sb.Append(firstWord.Substring(1).ToLower(CultureInfo.CurrentCulture));
                }
 
                return ToTitleCaseHelper(restOfString, sb.ToString());
            }
       }
 
        /// <summary>
        /// Returns the fieldname and switches from the given mergefield-instruction
        /// Note: the switches are always returned lowercase !
        /// Note 2: options holds values for formatting and text to insert before and/or after the field value.
        ///         options[0] = Formatting (Upper, Lower, Caps a.k.a. title case, FirstCap)
        ///         options[1] = Text to insert before data
        ///         options[2] = Text to insert after data
        /// </summary>
        /// <param name="field">The field being examined.</param>
        /// <param name="switches">An array of switches to apply to the field.</param>
        /// <param name="options">Formatting options to apply.</param>
        /// <returns>The name of the field.</returns>
        static string GetFieldNameWithOptions(SimpleField field, out string[] switches, out string[] options)
        {
            var a = field.GetAttribute("instr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            switches = new string[0];
            options = new string[3];
            string fieldname = string.Empty;
            string instruction = a.Value;
 
            if (!string.IsNullOrEmpty(instruction))
            {
                Match m = instructionRegEx.Match(instruction);
                if (m.Success)
                {
                    fieldname = m.Groups["name"].ToString().Trim();
                    options[0] = m.Groups["Format"].Value.Trim();
                    options[1] = m.Groups["PreText"].Value.Trim(); //becuase fixed we remove double quetes l.g
                    options[2] = m.Groups["PostText"].Value;//.Trim();1.0.0.8.3
                    int pos = fieldname.IndexOf('#');
                    if (!String.IsNullOrEmpty(options[1]))
                    {
                        var removeDoubleQuotesExpr = @"[^\\""]*(\\""[^\\""]*)*";
                        var removeDoubleQuotes = new Regex(removeDoubleQuotesExpr);
 
                        Match matchSyntax = removeDoubleQuotes.Match(options[1]);
 
                        if (matchSyntax.Success)
                            options[1] = matchSyntax.Groups[0] != null && String.IsNullOrEmpty(matchSyntax.Groups[0].Value) ? options[1] : matchSyntax.Groups[0].Value;
 
                    }
                    //start 1.0.0.8.3
                    if (!String.IsNullOrEmpty(options[2]))
                    {
                        var removeDoubleQuotesExpr = @"[^\\""]*(\\""[^\\""]*)*";
                        var removeDoubleQuotes = new Regex(removeDoubleQuotesExpr);
 
                        Match matchSyntax = removeDoubleQuotes.Match(options[2]);
 
                        if (matchSyntax.Success)
                            options[2] = matchSyntax.Groups[0] != null && String.IsNullOrEmpty(matchSyntax.Groups[0].Value) ? options[2] : matchSyntax.Groups[0].Value;
 
                    }
                    //end 1.0.0.8.3
                    if (pos > 0)
                    {
                        // Process the switches, correct the fieldname.
                        switches = fieldname.Substring(pos + 1).ToLower().Split(new char[] { '#' }, StringSplitOptions.RemoveEmptyEntries);
                        fieldname = fieldname.Substring(0, pos);
                    }
                }
            }
 
            return fieldname;
        }
 
        /// <summary>
        /// Executes the field switches on a given element.
        /// The possible switches are:
        /// <list>
        /// <li>dt : delete table</li>
        /// <li>dr : delete row</li>
        /// <li>dp : delete paragraph</li>
        /// </list>
        /// </summary>
        /// <param name="element">The element being operated on.</param>
        /// <param name="switches">The switched to be executed.</param>
        static void ExecuteSwitches(OpenXmlElement element, string[] switches)
        {
            if (switches == null || switches.Count() == 0)
            {
                return;
            }
 
            // check switches (switches are always lowercase)
            if (switches.Contains("dp"))
            {
                Paragraph p = GetFirstParent<Paragraph>(element);
                if (p != null)
                {
                    p.Remove();
                }
            }
            else if (switches.Contains("dr"))
            {
                TableRow row = GetFirstParent<TableRow>(element);
                if (row != null)
                {
                    row.Remove();
                }
            }
            else if (switches.Contains("dt"))
            {
                Table table = GetFirstParent<Table>(element);
                if (table != null)
                {
                    table.Remove();
                }
            }
        }
 
        /// <summary>
        /// Returns the first parent of a given <see cref="OpenXmlElement"/> that corresponds
        /// to the given type.
        /// This methods is different from the Ancestors-method on the OpenXmlElement in the sense that
        /// this method will return only the first-parent in direct line (closest to the given element).
        /// </summary>
        /// <typeparam name="T">The type of element being searched for.</typeparam>
        /// <param name="element">The element being examined.</param>
        /// <returns>The first parent of the element of the specified type.</returns>
        static T GetFirstParent<T>(OpenXmlElement element)
           where T : OpenXmlElement
        {
           if (element.Parent == null)
            {
                return null;
            }
            else if (element.Parent.GetType() == typeof(T))
            {
                return element.Parent as T;
            }
            else
            {
                return GetFirstParent<T>(element.Parent);
            }
        }
 
        /// <summary>
        /// Creates a paragraph to house text that should appear before or after the MergeField.
        /// </summary>
        /// <param name="text">The text to display.</param>
        /// <param name="fieldToMimic">The MergeField that will have its properties mimiced.</param>
        /// <returns>An OpenXml Paragraph ready to insert.</returns>
        static Paragraph GetPreOrPostParagraphToInsert(string text, SimpleField fieldToMimic)
        {
            Run runToInsert = GetRunElementForText(text, fieldToMimic);
            Paragraph paragraphToInsert = new Paragraph();
            paragraphToInsert.Append(runToInsert);
            return paragraphToInsert;
        }
 
        /// <summary>
        /// separte all words for handle  rtl (when english has than disable rtl) l.g
        /// Returns a <see cref="Run"/>-openxml element for the given text.
        /// Specific about this run-element is that it can describe multiple-line and tabbed-text.
        /// The <see cref="SimpleField"/> placeholder can be provided too, to allow duplicating the formatting.
        /// </summary>
        /// <param name="text">The text to be inserted.</param>
        /// <param name="placeHolder">The placeholder where the text will be inserted.</param>
        /// <param name="defaultHebFont">The default Hebrew Font where the text will be inserted 1.0.0.6.</param>
        /// <returns>A new <see cref="Run"/>-openxml element containing the specified text.</returns>
        internal static Run GetRunElementForTextRtlHandle(string text, SimpleField placeHolder, string defaultHebFont = "")
        {
            string rpr = null;
            if (placeHolder != null)
            {
                var xdoc = XDocument.Parse((placeHolder.Parent).OuterXml.Replace(placeHolder.OuterXml, string.Empty));
                if (xdoc.Root != null)
                {
                   var xrpr = xdoc.Root.Elements().FirstOrDefault(x => x.Name.LocalName == "rPr");
 
                    if (xrpr != null)
                        rpr = xrpr.ToString();
                }
            }
            var r = new Run();
            if (!string.IsNullOrEmpty(rpr))
            {
                r.AppendChild(new RunProperties(rpr));
            }
 
 
            if (!string.IsNullOrEmpty(text))
            {
                string[] split = text.Split(new string[] { "\n" }, StringSplitOptions.None);
                bool first = true;
                foreach (string s in split)
                {
                    if (!first)
                    {
                        r.Append(new Break());
                    }
 
                    first = false;
 
                    // then process tabs
                    bool firsttab = true;
                    string[] tabsplit = s.Split(new[] { "\t" }, StringSplitOptions.None);
                    foreach (string tabtext in tabsplit)
                    {
                        if (!firsttab)
                        {
                            r.Append(new TabChar());
                        }
                        // separte all words for handle  rtl (when english has than disable rtl)
                        //  string[] wordSpl = Regex.Split(tabtext, @"(?<=[\s+])");
                        string[] wordSpl = Regex.Split(tabtext, splitWhiteSpaceIncludeRegFormater);
 
                        foreach (var w in wordSpl)
                        {
                            Run rWord = new Run();
                            var txtTemp = new Text(w);
                            txtTemp.Space = SpaceProcessingModeValues.Preserve;
                            rWord.AppendChild(new RunProperties(rpr));
                            rWord.AppendChild<Text>(txtTemp);
                            r.AppendChild<Run>(rWord);
                            bool isHasEngChar = Regex.IsMatch(w, isEngCharcter);// @"[A-Za-z]+");  //l.g
                            var rprLtr = rWord.GetFirstChild<RunProperties>();
                            //1.0.0.7 remove rtl (false) on english chracters
                            //if (isHasEngChar && r.GetFirstChild<RunProperties>() != null)
                            //{
 
                            //    //if (rprLtr.RightToLeftText == null)
                            //    //    rprLtr.RightToLeftText = new RightToLeftText();
                            //    //rprLtr.RightToLeftText.Val = false;
 
                            //}
                            //else\
                            if (!isHasEngChar)
                            {
                                //1.0.0.5 save font hebrew
                                //set font hebrew (if not english)
                                rprLtr.AppendChild(GetHebrew());
                                //1.0.0.6
                                if (rprLtr.RunFonts == null && !String.IsNullOrEmpty(defaultHebFont))
                                {
                                    rprLtr.RunFonts = new RunFonts();
                                    //  rprLtr.RunFonts.Hint = FontTypeHintValues.EastAsia;
                                    rprLtr.RunFonts.Ascii = defaultHebFont;
                                }
 
                            }
                        }
                    }
                }
            }
            return r;
        }
        // 1.0.0.5
        static Languages GetHebrew()
        {
            Languages lang = new Languages();
            lang.Bidi = "he-IL";
            lang.Val = "he-IL";
            return lang;
        }
 
        /// <summary>
        /// Remove Paragraph from mail merge field for prevent line to be empty  l.g
        /// </summary>
        /// <param name="element">merge mail field</param>
        static void RemoveParentParagrph(OpenXmlElement element, IEnumerable<Paragraph> paragraphChange, ICollection<TableCell> tableCellRemoveParagraphs)
        {
            if (element == null)
                return;
            Paragraph paragraph = GetFirstParent<Paragraph>(element);
            if (paragraph != null)
            {
                if (paragraphChange.Contains(paragraph))
                    return;
                if (paragraph.Parent != null) //fixed 1.0.0.5
                {
 
                    //fixed 1.0.0.7 when has paragraph contain text do not remove him
                    if ((!paragraph.Descendants<Text>().Any()) || (!paragraph.Descendants<Text>().Where(t => t != null && !String.IsNullOrWhiteSpace(t.Text)).Any()))
                    {
                        //fixed 1.0.0.8
                        SetCellHasPargraphRemove(paragraph, tableCellRemoveParagraphs);
                        paragraph.Remove();
                    }
 
                }
            }
 
        }
        //fixed 1.0.0.8
        static void SetCellHasPargraphRemove(Paragraph paragraph, ICollection<TableCell> tableCellRemoveParagraphs)
        {
            if (paragraph.Parent != null && paragraph.Parent is TableCell)
            {
                if (!tableCellRemoveParagraphs.Contains((TableCell)paragraph.Parent))
                    tableCellRemoveParagraphs.Add((TableCell)paragraph.Parent);
            }
        }
        //fixed 1.0.0.8
        static void DelCellHasNoPargraphs(IEnumerable<TableCell> tableCellRemoveParagraphs)
        {
            foreach (var cell in tableCellRemoveParagraphs)
            {
                if (!cell.Descendants<Paragraph>().Any())
                {
                    if (cell.Parent != null)
                    {
                        if (cell.Parent is TableRow)
                        {    //when no has data but border has so keep it's as is row 1.0.0.8.1
                            //if (cell.Parent.Parent != null && cell.Parent.Parent is Table && cell.Parent.Parent.Descendants<TableBorders>().Any())
                            //{
                            if (cell.Parent.Parent != null && cell.Parent.Parent is Table)
                            {
                                Table tmpTable = (Table)cell.Parent.Parent;
                                TableRow firstRow = tmpTable.Elements<TableRow>().FirstOrDefault();
                                if (firstRow != null)
                                {
                                    if (firstRow.Descendants<Text>().Where(t => t != null && !String.IsNullOrWhiteSpace(t.Text)).Any())
                                    {
                                        // has header column on this table so do not remove cell only put empty text on the cell - 1.0.0.8.2
                                        cell.AppendChild(new Paragraph(new Run(new Text(String.Empty))));
                                        continue;
                                    }
                                }
                                //cell.AppendChild(new Paragraph(new Run(new Text(String.Empty))));
                                //continue;
                            }
                            //when no has data in all tablerow 
                            var isHasMoreCells = cell.Parent.Descendants<TableCell>().Count() > 1;
                            if (!isHasMoreCells)
                                if (cell.Parent.Parent != null)
                                {
                                    cell.Parent.Remove();
                                }
                        }
                    }
                    cell.Remove();
                }
            }
        }
    }
}
 
 

