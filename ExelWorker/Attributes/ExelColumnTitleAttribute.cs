using System;

namespace ExelWorker.Attributes
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ExelColumnTitleAttribute : Attribute
    {
        public string Title { get; set; }

        public ExelColumnTitleAttribute(string title)
        {
            Title = title;
        }
    }
}
