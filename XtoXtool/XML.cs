using System;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace XtoXtool
{
    // Примечание. Для запуска созданного кода может потребоваться NET Framework версии 4.5 или более поздней версии и .NET Core или Standard версии 2.0 или более поздней.
    /// <remarks/>
    [Serializable()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(AnonymousType = true)]
    [XmlRoot(Namespace = "", IsNullable = false)]
    public partial class Table
    {

        private TableRow[] rowsField;

        /// <remarks/>
        [XmlArrayItem("Row", IsNullable = false)]
        public TableRow[] Rows
        {
            get
            {
                return this.rowsField;
            }
            set
            {
                this.rowsField = value;
            }
        }
    }

    /// <remarks/>
    [Serializable()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(AnonymousType = true)]
    public partial class TableRow
    {

        private string wIC_NUMField;

        private uint wIC_CASE_NUMField;

        private System.DateTime wIC_DT_BEGINField;

        private System.DateTime wIC_DT_ENDField;

        private string wIC_STATUSField;

        private byte wIC_CDField;

        private string wIC_CD_NameField;

        private bool sIGN_ANLK_NARKOTIK_INTOXICATIONField;

        private bool vIOLATION_EXTENSIONField;

        private string hOSPITAL_NAMEField;

        private uint hOSPITAL_CODEField;

        private string nP_SURNAMEField;

        private string nP_NAMEField;

        private string nP_PATRONYMICField;

        private uint nP_NUMIDENTField;

        private string nP_DOC_NUMField;

        private byte nP_PDTField;

        /// <remarks/>
        public string WIC_NUM
        {
            get
            {
                return this.wIC_NUMField;
            }
            set
            {
                this.wIC_NUMField = value;
            }
        }

        /// <remarks/>
        public uint WIC_CASE_NUM
        {
            get
            {
                return this.wIC_CASE_NUMField;
            }
            set
            {
                this.wIC_CASE_NUMField = value;
            }
        }

        /// <remarks/>
        public System.DateTime WIC_DT_BEGIN
        {
            get
            {
                return this.wIC_DT_BEGINField;
            }
            set
            {
                this.wIC_DT_BEGINField = value;
            }
        }

        /// <remarks/>
        public System.DateTime WIC_DT_END
        {
            get
            {
                return this.wIC_DT_ENDField;
            }
            set
            {
                this.wIC_DT_ENDField = value;
            }
        }

        /// <remarks/>
        public string WIC_STATUS
        {
            get
            {
                return this.wIC_STATUSField;
            }
            set
            {
                this.wIC_STATUSField = value;
            }
        }

        /// <remarks/>
        public byte WIC_CD
        {
            get
            {
                return this.wIC_CDField;
            }
            set
            {
                this.wIC_CDField = value;
            }
        }

        /// <remarks/>
        public string WIC_CD_Name
        {
            get
            {
                return this.wIC_CD_NameField;
            }
            set
            {
                this.wIC_CD_NameField = value;
            }
        }

        /// <remarks/>
        public bool SIGN_ANLK_NARKOTIK_INTOXICATION
        {
            get
            {
                return this.sIGN_ANLK_NARKOTIK_INTOXICATIONField;
            }
            set
            {
                this.sIGN_ANLK_NARKOTIK_INTOXICATIONField = value;
            }
        }

        /// <remarks/>
        public bool VIOLATION_EXTENSION
        {
            get
            {
                return this.vIOLATION_EXTENSIONField;
            }
            set
            {
                this.vIOLATION_EXTENSIONField = value;
            }
        }

        /// <remarks/>
        public string HOSPITAL_NAME
        {
            get
            {
                return this.hOSPITAL_NAMEField;
            }
            set
            {
                this.hOSPITAL_NAMEField = value;
            }
        }

        /// <remarks/>
        public uint HOSPITAL_CODE
        {
            get
            {
                return this.hOSPITAL_CODEField;
            }
            set
            {
                this.hOSPITAL_CODEField = value;
            }
        }

        /// <remarks/>
        public string NP_SURNAME
        {
            get
            {
                return this.nP_SURNAMEField;
            }
            set
            {
                this.nP_SURNAMEField = value;
            }
        }

        /// <remarks/>
        public string NP_NAME
        {
            get
            {
                return this.nP_NAMEField;
            }
            set
            {
                this.nP_NAMEField = value;
            }
        }

        /// <remarks/>
        public string NP_PATRONYMIC
        {
            get
            {
                return this.nP_PATRONYMICField;
            }
            set
            {
                this.nP_PATRONYMICField = value;
            }
        }

        /// <remarks/>
        public uint NP_NUMIDENT
        {
            get
            {
                return this.nP_NUMIDENTField;
            }
            set
            {
                this.nP_NUMIDENTField = value;
            }
        }

        /// <remarks/>
        public string NP_DOC_NUM
        {
            get
            {
                return this.nP_DOC_NUMField;
            }
            set
            {
                this.nP_DOC_NUMField = value;
            }
        }

        /// <remarks/>
        public byte NP_PDT
        {
            get
            {
                return this.nP_PDTField;
            }
            set
            {
                this.nP_PDTField = value;
            }
        }
    }

    internal static class ParseHelpers
    {

        public static Stream ToStream(this string @this)
        {
            var stream = new MemoryStream();
            var writer = new StreamWriter(stream);
            writer.Write(@this);
            writer.Flush();
            stream.Position = 0;
            return stream;
        }

        public static T ParseXML<T>(this string @this) where T : class
        {
            var reader = XmlReader.Create(@this.Trim().ToStream(), new XmlReaderSettings() { ConformanceLevel = ConformanceLevel.Document });
            return new XmlSerializer(typeof(T)).Deserialize(reader) as T;
        }

    }

    public class Part : IEquatable<Part>
    {
        public int Tabn { get; set; }
        public string Icnum { get; set; }
        public string Ceh { get; set; }
        public override bool Equals(object obj)
        {
            if (obj == null) return false;
            Part objAsPart = obj as Part;
            if (objAsPart == null) return false;
            else return Equals(objAsPart);
        }
        public bool Equals(Part other)
        {
            if (other == null) return false;
            return (this.Icnum.Equals(other.Icnum));
        }
        public override int GetHashCode()
        {
            return Tabn;
        }
    }
}
