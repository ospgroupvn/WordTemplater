using System.Globalization;

namespace WordTemplater.Example;

internal class UpperCaseFirstLetter : IEvaluator
{
    public string Evaluate(object fieldValue, List<object> parameters)
    {
        if (fieldValue != null)
        {
            var strValue = fieldValue.ToString();
            var textInfo = new CultureInfo("en-US", false).TextInfo;
            return textInfo.ToTitleCase(strValue);
        }

        return string.Empty;
    }
}