using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Newtonsoft.Json.Linq;
using WP = DocumentFormat.OpenXml.Wordprocessing;

namespace WordTemplater;

public class WordTemplate : IDisposable
{
    private readonly List<RenderContext> _renderContexts;
    private readonly Stream _sourceStream;
    private readonly Dictionary<string, ITemplater> _templaterFactory;

    /// <summary>
    /// To init word template to export.
    /// </summary>
    /// <param name="sourceStream">
    /// The template word file stream.
    /// </param>
    /// <exception cref="ArgumentNullException"></exception>
    public WordTemplate(Stream sourceStream)
    {
        if (sourceStream == null)
        {
            throw new ArgumentNullException(nameof(sourceStream));
        }

        _sourceStream = sourceStream;
        _renderContexts = new List<RenderContext>();
        _templaterFactory = new Dictionary<string, ITemplater>();
        RegisterTemplater(string.Empty, new DefaultTemplater());
        RegisterTemplater(FunctionName.Sub, new SubTemplater());
        RegisterTemplater(FunctionName.Left, new LeftTemplater());
        RegisterTemplater(FunctionName.Right, new RightTemplater());
        RegisterTemplater(FunctionName.Trim, new TrimTemplater());
        RegisterTemplater(FunctionName.Upper, new UpperTemplater());
        RegisterTemplater(FunctionName.Lower, new LowerTemplater());
        RegisterTemplater(FunctionName.If, new IfTemplater());
        RegisterTemplater(FunctionName.Currency, new CurrencyTemplater());
        RegisterTemplater(FunctionName.Percentage, new PercentageTemplater());
        RegisterTemplater(FunctionName.Replace, new ReplaceTemplater());
        RegisterTemplater(FunctionName.BarCode, new BarCodeTemplater());
        RegisterTemplater(FunctionName.QRCode, new QRCodeTemplater());
        RegisterTemplater(FunctionName.Image, new ImageTemplater());
        RegisterTemplater(FunctionName.Html, new HtmlTemplater());
        RegisterTemplater(FunctionName.Word, new WordTemplater());
    }

    /// <summary>
    /// To register new format evaluator.
    /// </summary>
    /// <param name="name">
    /// The name of the format evaluator.
    /// </param>
    /// <param name="evaluator">
    /// An instance of IEvaluator.
    /// </param>
    public void RegisterEvaluator(string name, IEvaluator evaluator)
    {
        if (name != null)
        {
            name = name.ToLower();
            _templaterFactory[name] = new CustomTemplater(evaluator);
        }
    }

    private void RegisterTemplater(string name, ITemplater templater)
    {
        if (name != null)
        {
            name = name.ToLower();
            if (!_templaterFactory.ContainsKey(name))
            {
                _templaterFactory.Add(name, templater);
            }
            else
            {
                _templaterFactory[name] = templater;
            }
        }
    }

    /// <summary>
    /// To fill data into template file, then export it.
    /// </summary>
    /// <param name="data">
    /// The input data object. Such as JObject or any data model.
    /// </param>
    /// <param name="removeFallBack">
    /// Remove fall back element after exporting. Default value is true.
    /// </param>
    /// <returns>
    /// Return the exported file stream.
    /// </returns>
    public Stream Export(object data, bool removeFallBack = true)
    {
        if (data != null)
        {
            var targetStream = Utils.CloneStream(_sourceStream);
            if (targetStream != null)
            {
                using (var targetDocument = WordprocessingDocument.Open(targetStream, true))
                {
                    if (removeFallBack)
                    {
                        WordUtils.RemoveFallbackElements(targetDocument);
                    }

                    if (data is not JObject)
                    {
                        data = JObject.FromObject(data);
                    }

                    PrepareRenderContext(targetDocument);
                    RenderTemplate(_renderContexts.ToList(), data as JObject);
                    LoopTemplater.FillData(_renderContexts, data as JObject);
                }
            }

            targetStream.Position = 0;
            return targetStream;
        }

        return null;
    }


    private void PrepareRenderContext(WordprocessingDocument document)
    {
        var mainPart = document.MainDocumentPart;
        var documentPart = mainPart.Document;

        //find bookmark templates in header parts
        foreach (var headerPart in mainPart.HeaderParts)
        {
            _renderContexts.AddRange(PrepareRenderContext(headerPart.Header, headerPart));
        }

        //find bookmark templates in body parts
        _renderContexts.AddRange(PrepareRenderContext(documentPart.Body, mainPart));

        //find bookmark templates in footer parts
        foreach (var footerPart in mainPart.FooterParts)
        {
            _renderContexts.AddRange(PrepareRenderContext(footerPart.Footer, footerPart));
        }
    }

    private List<RenderContext> PrepareRenderContext(OpenXmlElement element, OpenXmlPart parentPart)
    {
        return PrepareRenderContext(new List<OpenXmlElement> { element }, parentPart);
    }

    private List<RenderContext> PrepareRenderContext(List<OpenXmlElement> elements, OpenXmlPart parentPart)
    {
        var rcRootList = new List<RenderContext>();
        var parents = new Stack<RenderContext>();
        var openContexts = new Stack<RenderContext>();
        var allMergeFieldNodes = new List<OpenXmlElement>();

        foreach (var element in elements)
        {
            allMergeFieldNodes.AddRange(element.Descendants().Where(x => IsMergeFieldNode(x)).ToList());
        }

        for (var i = 0; i < allMergeFieldNodes.Count; i++)
        {
            var mergeFieldNode = allMergeFieldNodes[i];
            var code = GetCode(mergeFieldNode);

            if (code.Contains(FunctionName.EndIf, StringComparison.OrdinalIgnoreCase))
            {
                if (openContexts.Count > 0)
                {
                    openContexts.Pop().MergeField.EndField = new MergeFieldTemplate(mergeFieldNode);
                }

                continue;
            }

            if (code.Contains(FunctionName.EndLoop, StringComparison.OrdinalIgnoreCase) || code.Contains(FunctionName.EndTable, StringComparison.OrdinalIgnoreCase))
            {
                if (parents.Count > 0)
                {
                    parents.Pop();
                }

                if (openContexts.Count > 0)
                {
                    openContexts.Pop().MergeField.EndField = new MergeFieldTemplate(mergeFieldNode);
                }

                continue;
            }

            var context = new RenderContext();
            context.MergeField = new MergeField();
            context.MergeField.ParentPart = parentPart;
            if (parents.Count > 0)
            {
                context.Parent = parents.Peek();
                context.Parent.ChildNodes.Add(context);
            }
            else
            {
                rcRootList.Add(context);
            }

            context.MergeField.StartField = new MergeFieldTemplate(mergeFieldNode);
            GetFormatTemplate(code, context);
            if (context.Templater != null)
            {
                var evaluator = context.Templater.Evaluator;
                if (evaluator != null)
                {
                    if (evaluator is LoopEvaluator)
                    {
                        parents.Push(context);
                    }
                    else if (evaluator is not ConditionEvaluator)
                    {
                        continue;
                    }

                    openContexts.Push(context);
                }
            }
        }

        return rcRootList;
    }

    private bool IsMergeFieldNode(OpenXmlElement x)
    {
        if (x is WP.SimpleField simpleField)
        {
            var fieldCode = simpleField.Instruction.Value;
            if (!string.IsNullOrWhiteSpace(fieldCode) && fieldCode.Contains(Constant.MERGEFORMAT))
            {
                return true;
            }
        }
        else if (x is WP.FieldCode code)
        {
            var fieldCode = code.Text;
            if (!string.IsNullOrWhiteSpace(fieldCode) && fieldCode.Contains(Constant.MERGEFORMAT))
            {
                return true;
            }
        }

        return false;
    }

    private string GetCode(OpenXmlElement node)
    {
        var fieldCode = string.Empty;
        if (node is WP.SimpleField field)
        {
            fieldCode = field.Instruction.Value;
        }
        else if (node is WP.FieldCode code)
        {
            fieldCode = code.Text;
        }

        return fieldCode.Trim();
    }


    private void GetFormatTemplate(string code, RenderContext context)
    {
        if (string.IsNullOrWhiteSpace(code))
        {
            return;
        }

        var i1 = code.IndexOf(" ");
        var i2 = code.IndexOf(Constant.MERGEFORMAT);
        if (i1 >= 0 && i2 >= 0 && i2 > i1)
        {
            code = code.Substring(i1, i2 - i1).Trim();
            if (code.StartsWith('"') && code.EndsWith('"'))
            {
                code = code.Substring(1);
                code = code.Substring(0, code.Length - 1);
            }

            var i3 = code.IndexOf('(');

            if (i3 > 0)
            {
                var fmtContent = code.Substring(0, i3);
                var i4 = fmtContent.IndexOf(':');
                var i5 = code.LastIndexOf(')');
                if (i5 < i3)
                {
                    i5 = code.Length - 1;
                }

                var paramContent = code.Substring(i3 + 1, i5 - i3 - 1).Trim();
                if (i4 > 0)
                {
                    context.FieldName = fmtContent.Substring(0, i4).Trim();
                    var function = fmtContent.Substring(i4 + 1).Trim().ToLower();
                    if (_templaterFactory.ContainsKey(function))
                    {
                        context.Templater = _templaterFactory[function];
                    }
                    else
                    {
                        context.Templater = new DefaultTemplater();
                    }

                    context.Parameters = paramContent;
                }
                else
                {
                    var function = fmtContent.Trim().ToLower();
                    if (function == FunctionName.If)
                    {
                        string op;
                        if (paramContent.IndexOf(OperatorName.Geq) > 0)
                        {
                            op = OperatorName.Geq;
                        }
                        else if (paramContent.IndexOf(OperatorName.Leq) > 0)
                        {
                            op = OperatorName.Leq;
                        }
                        else if (paramContent.IndexOf(OperatorName.Neq1) > 0)
                        {
                            op = OperatorName.Neq1;
                        }
                        else if (paramContent.IndexOf(OperatorName.Neq2) > 0)
                        {
                            op = OperatorName.Neq2;
                        }
                        else if (paramContent.IndexOf(OperatorName.Gt) > 0)
                        {
                            op = OperatorName.Gt;
                        }
                        else if (paramContent.IndexOf(OperatorName.Lt) > 0)
                        {
                            op = OperatorName.Lt;
                        }
                        else if (paramContent.IndexOf(OperatorName.Eq1) > 0)
                        {
                            op = OperatorName.Eq1;
                        }
                        else if (paramContent.IndexOf(OperatorName.Eq2) > 0)
                        {
                            op = OperatorName.Eq2;
                        }
                        else
                        {
                            return;
                        }

                        var sptContent = paramContent.Split(op);
                        context.FieldName = sptContent[0].Trim();
                        context.Parameters = sptContent[1].Trim();
                        context.Templater = new ConditionTemplater(op);
                    }
                    else if (function == FunctionName.Loop)
                    {
                        context.FieldName = paramContent;
                        context.Templater = new LoopTemplater();
                    }
                    else if (function == FunctionName.Table)
                    {
                        context.FieldName = paramContent;
                        context.Templater = new TableTemplater();
                    }
                }
            }
            else
            {
                context.Templater = new DefaultTemplater();
                var i6 = code.IndexOf(':');
                if (i6 > 0)
                {
                    context.FieldName = code.Substring(0, i6).Trim();
                    context.Parameters = code.Substring(i6 + 1).Trim();
                }
                else
                {
                    context.FieldName = code;
                }
            }
        }
    }

    private void RenderTemplate(List<RenderContext> renderContexts, JObject dataObj)
    {
        if (dataObj != null)
        {
            foreach (var rc in renderContexts)
            {
                var value = dataObj.GetValue(rc.FieldName, StringComparison.OrdinalIgnoreCase);
                if (rc.Templater != null && rc.Templater.Evaluator is LoopEvaluator && value is JArray array && rc.MergeField.StartField != null && rc.MergeField.EndField != null)
                {
                    RenderTemplate(rc, array);
                }
            }
        }
    }

    private void RenderTemplate(RenderContext rc, JArray arrData)
    {
        var template = GetRepeatingTemplate(rc);
        if (arrData.Count > 1)
        {
            var generatedNodes = new List<OpenXmlElement>();
            generatedNodes.AddRange(template.TemplateElements);

            for (var i = 1; i < arrData.Count; i++)
            {
                generatedNodes.AddRange(template.CloneAndAppendTemplate());
            }

            var allContexts = PrepareRenderContext(generatedNodes, rc.MergeField.ParentPart);

            var firstRC = Find(allContexts, rc);
            if (firstRC != null)
            {
                if (firstRC.Parent != null)
                {
                    allContexts = firstRC.Parent.ChildNodes;
                }

                var firstIndex = allContexts.IndexOf(firstRC);
                if (firstIndex >= 0)
                {
                    var lastRC = rc;
                    for (var j = 0; j < arrData.Count; j++)
                    {
                        var ct = allContexts[j + firstIndex];
                        ct.Index = j;
                        if (rc.Parent != null)
                        {
                            ct.Parent = rc.Parent;
                            var pos = rc.Parent.ChildNodes.IndexOf(lastRC);
                            if (pos > -1)
                            {
                                if (pos < rc.Parent.ChildNodes.Count - 1)
                                {
                                    rc.Parent.ChildNodes.Insert(pos + 1, ct);
                                }
                                else
                                {
                                    rc.Parent.ChildNodes.Add(ct);
                                }

                                lastRC = ct;
                            }
                        }
                        else if (rc.Parent == null)
                        {
                            var pos = _renderContexts.IndexOf(lastRC);
                            if (pos > -1)
                            {
                                if (pos < _renderContexts.Count - 1)
                                {
                                    _renderContexts.Insert(pos + 1, ct);
                                }
                                else
                                {
                                    _renderContexts.Add(ct);
                                }

                                lastRC = ct;
                            }
                        }

                        var value = arrData[j];
                        if (ct.ChildNodes.Count > 0)
                        {
                            RenderTemplate(ct.ChildNodes.ToList(), value as JObject);
                        }
                    }

                    if (rc.Parent == null)
                    {
                        _renderContexts.Remove(rc);
                    }
                    else
                    {
                        rc.Parent.ChildNodes.Remove(rc);
                        rc.Parent = null;
                    }
                }
            }
        }
        else if (arrData.Count == 0)
        {
            foreach (var el in template.TemplateElements)
            {
                el.Remove();
            }
        }
    }

    private RepeatingTemplate GetRepeatingTemplate(RenderContext context)
    {
        var mfTemplate = new RepeatingTemplate();

        //find the ascendant of both start and end bookmark node
        var mfStart = context.MergeField.StartField.StartNode;
        var parentNode = mfStart;

        while (parentNode.Parent != null && !parentNode.Parent.Descendants().Any(el => el == context.MergeField.EndField.EndNode))
        {
            parentNode = parentNode.Parent;
        }

        if (parentNode != null)
        {
            OpenXmlElement lastChildNode = null;
            OpenXmlElement startNode = parentNode, endNode = null;

            if (context.Templater.Evaluator is TableEvaluator)
            {
                var tableParentNode = startNode;
                while (tableParentNode != null)
                {
                    if (tableParentNode is WP.Table || tableParentNode.Descendants<WP.Table>().Any()
                                                    || tableParentNode == context.MergeField.EndField.EndNode
                                                    || tableParentNode.Descendants().Any(el => context.MergeField.EndField.EndNode == el))
                    {
                        break;
                    }

                    tableParentNode = tableParentNode.NextSibling();
                }

                WP.Table tableNode = null;
                if (tableParentNode is WP.Table node)
                {
                    tableNode = node;
                }
                else
                {
                    tableNode = tableParentNode.Descendants<WP.Table>().FirstOrDefault();
                }

                if (tableNode != null)
                {
                    WP.TableCell firstCell = null, lastCell = null;
                    foreach (var row in tableNode.ChildElements.Where(r => r is WP.TableRow && r.Descendants().Any(x => IsMergeFieldNode(x))))
                    {
                        if (firstCell == null)
                        {
                            firstCell = ((WP.TableRow)row).Descendants<WP.TableCell>().FirstOrDefault();
                        }

                        mfTemplate.TemplateElements.Add(row);
                        lastChildNode = row;
                    }

                    lastCell = ((WP.TableRow)lastChildNode).Descendants<WP.TableCell>().LastOrDefault();

                    MoveMergeFieldTo(context.MergeField.StartField, firstCell);
                    MoveMergeFieldTo(context.MergeField.EndField, lastCell, false);
                }
            }
            else
            {
                var curentNode = startNode;
                while (curentNode != null)
                {
                    if (!mfTemplate.TemplateElements.Contains(curentNode))
                    {
                        mfTemplate.TemplateElements.Add(curentNode);
                        lastChildNode = curentNode;
                    }

                    if (curentNode == context.MergeField.EndField.EndNode || curentNode.Descendants().Any(el => context.MergeField.EndField.EndNode == el))
                    {
                        endNode = curentNode;
                        break;
                    }

                    curentNode = curentNode.NextSibling();
                }
            }

            if (lastChildNode != null)
            {
                if (lastChildNode.NextSibling() != null)
                {
                    mfTemplate.LastNode = lastChildNode.NextSibling();
                }

                mfTemplate.ParentNode = lastChildNode.Parent;
            }
        }

        return mfTemplate;
    }

    private void MoveMergeFieldTo(MergeFieldTemplate fieldTemplate, WP.TableCell? tableCell, bool forBeginning = true)
    {
        if (fieldTemplate != null && tableCell != null)
        {
            var prg = tableCell.Descendants<WP.Paragraph>().FirstOrDefault();
            if (prg == null)
            {
                prg = new WP.Paragraph();
                tableCell.Append(prg);
            }

            var parentParagraph = fieldTemplate.StartNode.Ancestors<WP.Paragraph>().FirstOrDefault();
            if (forBeginning)
            {
                foreach (var item in fieldTemplate.GetAllElements(true))
                {
                    item.Remove();
                    prg.InsertAt(item, 0);
                }
            }
            else
            {
                foreach (var item in fieldTemplate.GetAllElements())
                {
                    item.Remove();
                    prg.Append(item);
                }
            }

            if (parentParagraph != null && parentParagraph.Parent != null && parentParagraph.InnerText.Trim() == string.Empty)
            {
                parentParagraph.Remove();
            }
        }
    }

    private RenderContext Find(List<RenderContext> renderContexts, RenderContext renderContext)
    {
        foreach (var rc in renderContexts)
        {
            if (rc.MergeField.StartField.StartNode == renderContext.MergeField.StartField.StartNode)
            {
                return rc;
            }

            if (rc.ChildNodes.Count > 0)
            {
                var rmf = Find(rc.ChildNodes, renderContext);
                if (rmf != null)
                {
                    return rmf;
                }
            }
        }

        return null;
    }

    public void Dispose()
    {
        if (_sourceStream != null)
        {
            _sourceStream.Dispose();
        }
    }
}