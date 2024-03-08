using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PdfMerge.Model
{
    public class PdfElement
    {
        private string fullpath;
        private string shortpath;

        public PdfElement()
        {
        }

        public PdfElement(string fullpath, string shortpath)
        {
            this.fullpath = fullpath;
            this.shortpath = shortpath;
        }

        public override string? ToString()
        {
            return base.ToString();
        }

        public string Fullpath
        {
            get => fullpath;
            set => fullpath = value;
        }
        public string Shortpath
        {
            get => shortpath;
            set => shortpath = value;
        }


    }
}
