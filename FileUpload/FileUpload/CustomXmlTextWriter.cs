using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace FileUpload
{
    public class CustomXmlTextWriter : XmlTextWriter
    {
        public CustomXmlTextWriter(string fileName)
        : base(fileName, Encoding.UTF8)
        {
            this.Formatting = Formatting.Indented;
        }

        public override void WriteEndElement()
        {
            this.WriteFullEndElement();
        }
    }
}
