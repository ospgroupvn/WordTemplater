using System.Diagnostics;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using Newtonsoft.Json.Linq;
using WP = DocumentFormat.OpenXml.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using DRAW = DocumentFormat.OpenXml.Drawing;

namespace WordTemplater;

internal interface ITemplater
{
    /// <summary>
    /// Indicates the evaluator <see cref="IEvaluator"/>
    /// </summary>
    IEvaluator Evaluator { get; }

    /// <summary>
    /// Replace the template with the corresponding data.
    /// </summary>
    /// <param name="value">The input value</param>
    /// <param name="context">The render context</param>
    void FillData(JToken? value, RenderContext context);
}

internal class Templater : ITemplater
{
    private protected readonly IEvaluator _evaluator;
    IEvaluator ITemplater.Evaluator => _evaluator;

    internal Templater(IEvaluator evaluator)
    {
        _evaluator = evaluator;
    }

    void ITemplater.FillData(JToken? value, RenderContext context)
    {
        if (value is JValue jValue)
        {
            var jvalue = jValue.Value;
            if (jvalue != null)
            {
                var eval = "";
                var templater = context.Templater;
                if (templater != null)
                {
                    try
                    {
                        eval = templater.Evaluator.Evaluate(jvalue, templater.Evaluator is DefaultEvaluator ? new List<object> { context.Parameters } : Utils.PaserParametters(context.Parameters));
                    }
                    catch
                    {
                        eval = jvalue.ToString();
                    }
                }
                else
                {
                    eval = jvalue.ToString();
                }

                var textNode = context.MergeField.StartField?.RemoveAllExceptTextNode();
                if (textNode != null)
                {
                    textNode.Space = SpaceProcessingModeValues.Preserve;
                    textNode.Text = eval;
                }
            }
            else
            {
                context.MergeField.StartField?.RemoveAll();
            }
        }
    }
}

[DebuggerDisplay("[Custom Templater]")]
internal class CustomTemplater : Templater, ITemplater
{
    internal CustomTemplater(IEvaluator evaluator) : base(evaluator) { }
}

[DebuggerDisplay("[Default Templater]")]
internal class DefaultTemplater : Templater, ITemplater
{
    internal DefaultTemplater() : base(new DefaultEvaluator()) { }
}

[DebuggerDisplay("[Sub Templater]")]
internal class SubTemplater : Templater, ITemplater
{
    internal SubTemplater() : base(new SubEvaluator()) { }
}

[DebuggerDisplay("[Left Templater]")]
internal class LeftTemplater : Templater, ITemplater
{
    internal LeftTemplater() : base(new LeftEvaluator()) { }
}

[DebuggerDisplay("[Right Templater]")]
internal class RightTemplater : Templater, ITemplater
{
    internal RightTemplater() : base(new RightEvaluator()) { }
}

[DebuggerDisplay("[Trim Templater]")]
internal class TrimTemplater : Templater, ITemplater
{
    internal TrimTemplater() : base(new TrimEvaluator()) { }
}

[DebuggerDisplay("[Upper Templater]")]
internal class UpperTemplater : Templater, ITemplater
{
    internal UpperTemplater() : base(new UpperEvaluator()) { }
}

[DebuggerDisplay("[Lower Templater]")]
internal class LowerTemplater : Templater, ITemplater
{
    internal LowerTemplater() : base(new LowerEvaluator()) { }
}

[DebuggerDisplay("[Currency Templater]")]
internal class CurrencyTemplater : Templater, ITemplater
{
    internal CurrencyTemplater() : base(new CurrencyEvaluator()) { }
}

[DebuggerDisplay("[Percentage Templater]")]
internal class PercentageTemplater : Templater, ITemplater
{
    internal PercentageTemplater() : base(new PercentageEvaluator()) { }
}

[DebuggerDisplay("[Replace Templater]")]
internal class ReplaceTemplater : Templater, ITemplater
{
    internal ReplaceTemplater() : base(new ReplaceEvaluator()) { }
}

[DebuggerDisplay("[If Templater]")]
internal class IfTemplater : Templater, ITemplater
{
    internal IfTemplater() : base(new IfEvaluator()) { }
}

[DebuggerDisplay("[Condition Templater]")]
internal class ConditionTemplater : Templater, ITemplater
{
    internal string _operator { get; set; }

    internal ConditionTemplater(string op) : base(new ConditionEvaluator())
    {
        _operator = op;
    }

    void ITemplater.FillData(JToken? value, RenderContext context)
    {
        if (value is JValue jValue)
        {
            var jvalue = jValue.Value;
            var listParam = Utils.PaserParametters(context.Parameters);
            listParam.Insert(0, _operator);
            var eval = _evaluator.Evaluate(jvalue, listParam);
            if (string.Compare(true.ToString(), eval, StringComparison.OrdinalIgnoreCase) == 0)
            {
                context.MergeField.StartField?.RemoveAll(true);
                context.MergeField.EndField?.RemoveAll(true);
            }
            else
            {
                var start = context.MergeField.StartField.GetAllElements()[0];
                var end = context.MergeField.EndField.GetAllElements(true)[0];
                WordUtils.RemoveFromNodeToNode(start, end);
            }
        }
    }
}

[DebuggerDisplay("[Loop Templater]")]
internal class LoopTemplater : Templater, ITemplater
{
    internal LoopTemplater() : base(new LoopEvaluator()) { }

    internal LoopTemplater(IEvaluator evaluator) : base(evaluator) { }

    void ITemplater.FillData(JToken? value, RenderContext context)
    {
        if (value is JArray arr)
        {
            if (arr.Count > 0)
            {
                var arrItem = arr[context.Index];
                if (arrItem is JObject arrItemJob)
                {
                    arrItemJob[Constant.CURRENT_INDEX] = context.Index + 1;
                    arrItemJob[Constant.IS_LAST] = (context.Index == arr.Count - 1);
                    FillData(context.ChildNodes, arrItemJob);
                }
                else if (arrItem is JValue item)
                {
                    var jval = new JObject();
                    jval[Constant.CURRENT_NODE] = item;
                    jval[Constant.CURRENT_INDEX] = context.Index + 1;
                    jval[Constant.IS_LAST] = (context.Index == arr.Count - 1);
                    FillData(context.ChildNodes, jval);
                }

                context.MergeField.StartField?.RemoveAll(true);
                context.MergeField.EndField?.RemoveAll(true);
            }
        }
    }

    internal static void FillData(List<RenderContext> renderContexts, JObject data)
    {
        foreach (var context in renderContexts)
        {
            var value = data.GetValue(context.FieldName, StringComparison.OrdinalIgnoreCase);
            context.Templater?.FillData(value, context);
        }
    }
}

[DebuggerDisplay("[Table Templater]")]
internal class TableTemplater : LoopTemplater, ITemplater
{
    internal TableTemplater() : base(new TableEvaluator()) { }
}

[DebuggerDisplay("[Image Templater]")]
internal class ImageTemplater : Templater, ITemplater
{
    internal ImageTemplater() : base(new ImageEvaluator()) { }

    internal ImageTemplater(IEvaluator evaluator) : base(evaluator) { }

    void ITemplater.FillData(JToken? value, RenderContext context)
    {
        if (value is JValue jValue)
        {
            var jvalue = jValue.Value;
            if (jvalue != null)
            {
                var base64Img = _evaluator.Evaluate(jvalue.ToString(), null);
                var stream = new MemoryStream(Convert.FromBase64String(base64Img));

                var drawing = context.MergeField.StartField.StartNode.Ancestors<WP.Drawing>().FirstOrDefault();
                OpenXmlElement imageParentElement = null;
                Size displaySize = null;
                DRAW.SourceRectangle sourceRectangle = null;
                var shapeType = DRAW.ShapeTypeValues.Rectangle;
                if (drawing != null)
                {
                    var run = drawing.Ancestors<WP.Run>().FirstOrDefault();
                    DRAW.GraphicData graphicData = null;
                    displaySize = GetShapeSize(drawing.Descendants<DRAW.Extents>().FirstOrDefault());
                    if (displaySize == null)
                    {
                        displaySize = WordUtils.GetImageSize(stream);
                    }

                    if (displaySize == null)
                    {
                        goto DONE;
                    }

                    graphicData = drawing.Descendants<DRAW.GraphicData>().FirstOrDefault();
                    if (graphicData != null)
                    {
                        var geometry = graphicData.Descendants<DRAW.PresetGeometry>().FirstOrDefault();
                        if (geometry != null && geometry.Preset.HasValue)
                        {
                            shapeType = geometry.Preset.Value;
                        }

                        graphicData.RemoveAllChildren();
                        graphicData.Uri = Constant.PICTURE_NAMESPACE;
                        imageParentElement = graphicData;
                    }
                    else
                    {
                        if (run == null)
                        {
                            goto DONE;
                        }

                        imageParentElement = run;
                    }

                    var originalSize = WordUtils.GetImageSize(stream);
                    var ratio = Math.Max((double)displaySize.Width / originalSize.Width, (double)displaySize.Height / originalSize.Height);
                    var newImageSize = new Size((long)(originalSize.Width * ratio), (long)(originalSize.Height * ratio));
                    int percentVertical = WordUtils.ToThousandPercent((double)(newImageSize.Height - displaySize.Height) / 2 / newImageSize.Height * 100),
                        percentHorizontal = WordUtils.ToThousandPercent((double)(newImageSize.Width - displaySize.Width) / 2 / newImageSize.Width * 100);
                    sourceRectangle = new DRAW.SourceRectangle();
                    if (percentHorizontal > 0)
                    {
                        sourceRectangle.Left = sourceRectangle.Right = percentHorizontal;
                    }

                    if (percentVertical > 0)
                    {
                        sourceRectangle.Top = sourceRectangle.Bottom = percentVertical;
                    }
                }
                else
                {
                    var originalSize = WordUtils.GetImageSize(stream);
                    var percent = GetPercent(context.Parameters);
                    if (percent.HasValue)
                    {
                        (var width, var height) = WordUtils.GetPageSize(context.MergeField.StartField.StartNode);
                        double imageWidth = width * percent.Value, imageHeight = height * percent.Value;
                        var ratio = Math.Min(imageWidth / originalSize.Width, imageHeight / originalSize.Height);
                        displaySize = new Size((long)(originalSize.Width * ratio), (long)(originalSize.Height * ratio));
                    }
                    else
                    {
                        displaySize = originalSize;
                    }

                    (var drawingElement, var graphicDataElement) = CreateNewGraphicElement(displaySize);
                    imageParentElement = graphicDataElement;
                    context.MergeField.StartField.RootNode.InsertBeforeSelf(drawingElement);
                }

                var imgElement = CreateNewPictureElement(stream, context.MergeField.ParentPart, displaySize, shapeType, sourceRectangle);
                imageParentElement.Append(imgElement);
            }
        }

        DONE:
        context.MergeField.StartField.RemoveAll();
    }

    private ImagePart AddImagePart(OpenXmlPart parentPart)
    {
        switch (parentPart)
        {
            case HeaderPart headerPart:
                return headerPart.AddImagePart(ImagePartType.Png);
            case FooterPart footerPart:
                return footerPart.AddImagePart(ImagePartType.Png);
        }

        return ((MainDocumentPart)parentPart).AddImagePart(ImagePartType.Png);
    }

    private (OpenXmlElement drawingElement, OpenXmlElement graphicDataElement) CreateNewGraphicElement(Size size)
    {
        var name = Guid.NewGuid().ToString();
        var graphicDataElement = new DRAW.GraphicData();
        graphicDataElement.Uri = Constant.PICTURE_NAMESPACE;
        var element =
            new WP.Drawing(
                new Inline(
                    new Extent { Cx = WordUtils.PixelToEmu(size.Width), Cy = WordUtils.PixelToEmu(size.Height) },
                    new EffectExtent
                    {
                        LeftEdge = 0L,
                        TopEdge = 0L,
                        RightEdge = 0L,
                        BottomEdge = 0L
                    },
                    new DocProperties
                    {
                        Id = Utils.GetUintId(),
                        Name = name
                    },
                    new DRAW.NonVisualGraphicFrameDrawingProperties(
                        new DRAW.GraphicFrameLocks { NoChangeAspect = true }),
                    new DRAW.Graphic(graphicDataElement))
                {
                    DistanceFromTop = 0U,
                    DistanceFromBottom = 0U,
                    DistanceFromLeft = 0U,
                    DistanceFromRight = 0U,
                    EditId = Utils.GetRandomHexNumber(8)
                });
        return (element, graphicDataElement);
    }

    private PIC.Picture CreateNewPictureElement(MemoryStream stream, OpenXmlPart parentPart, Size size, DRAW.ShapeTypeValues type, DRAW.SourceRectangle sourceRect = null)
    {
        var imagePart = AddImagePart(parentPart);
        var imageId = parentPart.GetIdOfPart(imagePart);
        imagePart.FeedData(stream);
        var blipFill = new PIC.BlipFill(
            new DRAW.Blip(
                new DRAW.BlipExtensionList())
            {
                CompressionState = DRAW.BlipCompressionValues.Print,
                Embed = imageId
            })
        {
            RotateWithShape = true
        };
        if (sourceRect != null)
        {
            blipFill.Append(sourceRect);
        }

        blipFill.Append(new DRAW.Stretch());
        var element = new PIC.Picture(
            new PIC.NonVisualPictureProperties(
                new PIC.NonVisualDrawingProperties
                {
                    Id = Utils.GetUintId(),
                    Name = string.Format(Constant.DEFAULT_IMAGE_FILE_NAME, Guid.NewGuid().ToString())
                },
                new PIC.NonVisualPictureDrawingProperties()),
            blipFill,
            new PIC.ShapeProperties(
                new DRAW.Transform2D(
                    new DRAW.Offset { X = 0L, Y = 0L },
                    new DRAW.Extents { Cx = WordUtils.PixelToEmu(size.Width), Cy = WordUtils.PixelToEmu(size.Height) }),
                new DRAW.PresetGeometry(
                    new DRAW.AdjustValueList())
                {
                    Preset = type
                }));
        return element;
    }

    private Size GetShapeSize(DRAW.Extents extents)
    {
        if (extents != null)
        {
            var w = extents.Cx;
            var h = extents.Cy;
            if (w.HasValue && h.HasValue)
            {
                return new Size((long)WordUtils.EmuToPixels(w.Value), (long)WordUtils.EmuToPixels(h.Value));
            }
        }

        return null;
    }

    private double? GetPercent(string parameters)
    {
        if (parameters == null)
        {
            return null;
        }

        parameters = parameters.Trim();
        if (parameters.Length == 0)
        {
            return null;
        }

        return Utils.GetDouble(parameters);
    }

    private DRAW.ShapeTypeValues GetEnumShapeType(string shapeType)
    {
        switch (shapeType)
        {
            case "line": return DRAW.ShapeTypeValues.Line;
            case "lineInv": return DRAW.ShapeTypeValues.LineInverse;
            case "triangle": return DRAW.ShapeTypeValues.Triangle;
            case "rtTriangle": return DRAW.ShapeTypeValues.RightTriangle;
            case "rect": return DRAW.ShapeTypeValues.Rectangle;
            case "diamond": return DRAW.ShapeTypeValues.Diamond;
            case "parallelogram": return DRAW.ShapeTypeValues.Parallelogram;
            case "trapezoid": return DRAW.ShapeTypeValues.Trapezoid;
            case "nonIsoscelesTrapezoid": return DRAW.ShapeTypeValues.NonIsoscelesTrapezoid;
            case "pentagon": return DRAW.ShapeTypeValues.Pentagon;
            case "hexagon": return DRAW.ShapeTypeValues.Hexagon;
            case "heptagon": return DRAW.ShapeTypeValues.Heptagon;
            case "octagon": return DRAW.ShapeTypeValues.Octagon;
            case "decagon": return DRAW.ShapeTypeValues.Decagon;
            case "dodecagon": return DRAW.ShapeTypeValues.Dodecagon;
            case "star4": return DRAW.ShapeTypeValues.Star4;
            case "star5": return DRAW.ShapeTypeValues.Star5;
            case "star6": return DRAW.ShapeTypeValues.Star6;
            case "star7": return DRAW.ShapeTypeValues.Star7;
            case "star8": return DRAW.ShapeTypeValues.Star8;
            case "star10": return DRAW.ShapeTypeValues.Star10;
            case "star12": return DRAW.ShapeTypeValues.Star12;
            case "star16": return DRAW.ShapeTypeValues.Star16;
            case "star24": return DRAW.ShapeTypeValues.Star24;
            case "star32": return DRAW.ShapeTypeValues.Star32;
            case "roundRect": return DRAW.ShapeTypeValues.RoundRectangle;
            case "round1Rect": return DRAW.ShapeTypeValues.Round1Rectangle;
            case "round2SameRect": return DRAW.ShapeTypeValues.Round2SameRectangle;
            case "round2DiagRect": return DRAW.ShapeTypeValues.Round2DiagonalRectangle;
            case "snipRoundRect": return DRAW.ShapeTypeValues.SnipRoundRectangle;
            case "snip1Rect": return DRAW.ShapeTypeValues.Snip1Rectangle;
            case "snip2SameRect": return DRAW.ShapeTypeValues.Snip2SameRectangle;
            case "snip2DiagRect": return DRAW.ShapeTypeValues.Snip2DiagonalRectangle;
            case "plaque": return DRAW.ShapeTypeValues.Plaque;
            case "ellipse": return DRAW.ShapeTypeValues.Ellipse;
            case "teardrop": return DRAW.ShapeTypeValues.Teardrop;
            case "homePlate": return DRAW.ShapeTypeValues.HomePlate;
            case "chevron": return DRAW.ShapeTypeValues.Chevron;
            case "pieWedge": return DRAW.ShapeTypeValues.PieWedge;
            case "pie": return DRAW.ShapeTypeValues.Pie;
            default: return DRAW.ShapeTypeValues.Rectangle;
        }
    }
}

[DebuggerDisplay("[BarCode Templater]")]
internal class BarCodeTemplater : ImageTemplater, ITemplater
{
    internal BarCodeTemplater() : base(new BarCodeEvaluator()) { }
}

internal class QRCodeTemplater : ImageTemplater, ITemplater
{
    internal QRCodeTemplater() : base(new QRCodeEvaluator()) { }
}

[DebuggerDisplay("[Html Templater]")]
internal class HtmlTemplater : Templater, ITemplater
{
    internal HtmlTemplater() : base(new HtmlEvaluator()) { }

    void ITemplater.FillData(JToken? value, RenderContext context)
    {
        var isRemovedMergeField = false;
        if (value is JValue jValue)
        {
            var jvalue = jValue.Value;
            var eval = _evaluator.Evaluate(jvalue, null);
            var stream = new MemoryStream(Encoding.UTF8.GetBytes(string.Format(Constant.HTML_PATTERN, eval)));
            AlternativeFormatImportPart formatImportPart = null;
            if (context.MergeField.ParentPart is MainDocumentPart part)
            {
                formatImportPart = part.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html);
            }
            else if (context.MergeField.ParentPart is HeaderPart headerPart)
            {
                formatImportPart = headerPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html);
            }
            else if (context.MergeField.ParentPart is FooterPart footerPart)
            {
                formatImportPart = footerPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html);
            }

            if (formatImportPart != null)
            {
                formatImportPart.FeedData(stream);
                var altChunk = new WP.AltChunk();
                altChunk.Id = context.MergeField.ParentPart.GetIdOfPart(formatImportPart);
                var node = context.MergeField.StartField?.RemoveAllExceptTextNode();
                if (node != null && node.Parent != null && node.Parent.Parent != null)
                {
                    node.Parent.InsertAfterSelf(new WP.Run(altChunk));
                    node.Parent.Remove();
                    isRemovedMergeField = true;
                }
            }

            stream.Dispose();
        }

        if (!isRemovedMergeField)
        {
            context.MergeField.StartField?.RemoveAll();
        }
    }
}

[DebuggerDisplay("[Word Templater]")]
internal class WordTemplater : Templater, ITemplater
{
    internal WordTemplater() : base(new WordEvaluator()) { }

    void ITemplater.FillData(JToken? value, RenderContext context)
    {
        if (value is JValue jValue)
        {
            if (context.MergeField.ParentPart is MainDocumentPart part)
            {
                var body = part.Document.Body;
                var startNodeToInsert = context.MergeField.StartField.StartNode.Ancestors().FirstOrDefault(a => a.Parent == body);
                if (startNodeToInsert != null)
                {
                    var jvalue = jValue.Value;
                    var eval = _evaluator.Evaluate(jvalue, null);
                    Stream stream = new MemoryStream(Convert.FromBase64String(eval));
                    var wordDocument = WordprocessingDocument.Open(stream, false);
                    var mappingRID = new Dictionary<string, string>();
                    foreach (var p in wordDocument.MainDocumentPart.Parts)
                    {
                        //ignore header and footer data
                        if (p.OpenXmlPart is HeaderPart or FooterPart)
                        {
                            continue;
                        }

                        try
                        {
                            var rId = Utils.GetUniqueStringID();
                            context.MergeField.ParentPart.AddPart(p.OpenXmlPart, rId);
                            mappingRID.Add(p.RelationshipId, rId);
                        }
                        catch { }
                    }

                    foreach (var el in wordDocument.MainDocumentPart.Document.Body.Elements())
                    {
                        if (el is WP.SectionProperties)
                        {
                            continue;
                        }

                        var newEl = el.CloneNode(true);
                        var subEls = newEl.Descendants().Where(x => { return x.GetAttributes().Where(a => mappingRID.ContainsKey(a.Value)).FirstOrDefault().LocalName != null; });

                        foreach (var x in subEls)
                        {
                            var att = x.GetAttributes().Where(a => mappingRID.ContainsKey(a.Value)).FirstOrDefault();
                            var newAttr = new OpenXmlAttribute(att.Prefix, att.LocalName, att.NamespaceUri, mappingRID[att.Value]);
                            x.SetAttribute(newAttr);
                        }

                        if (startNodeToInsert is WP.Paragraph && newEl is WP.Paragraph)
                        {
                            var paraProp = startNodeToInsert.Elements<WP.ParagraphProperties>().FirstOrDefault();
                            var oldParaProp = newEl.Elements<WP.ParagraphProperties>().FirstOrDefault();
                            if (paraProp != null)
                            {
                                var newParaProp = new WP.ParagraphProperties();
                                if (oldParaProp != null)
                                {
                                    foreach (var item in oldParaProp.ChildElements)
                                    {
                                        if (!(item is WP.Indentation))
                                        {
                                            newParaProp.Append(item.CloneNode(true));
                                        }
                                    }

                                    oldParaProp.Remove();
                                }

                                foreach (var item in paraProp.ChildElements)
                                {
                                    if (!newParaProp.Elements().Any(x => x.GetType() == item.GetType()))
                                    {
                                        newParaProp.Append(item.CloneNode(true));
                                    }
                                }

                                newEl.InsertAt(newParaProp, 0);
                            }
                        }

                        startNodeToInsert.InsertBeforeSelf(newEl);
                    }

                    wordDocument.Dispose();
                }
            }
        }

        context.MergeField.StartField?.RemoveAll(true);
    }
}