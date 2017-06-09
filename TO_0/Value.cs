using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace TO_0
{
    internal class Value<T> : EnumValue<BorderValues>
    {
        public Value(BorderValues value) : base(value)
        {
        }
    }
}