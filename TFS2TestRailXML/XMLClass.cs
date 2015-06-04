using System.Xml.Serialization;

namespace TFS2TestRailXML
{
    internal class XMLClass
    {
        /// <remarks />
        [XmlType(AnonymousType = true)]
        [XmlRoot(Namespace = "", IsNullable = false)]
        public class Suite
        {
            /// <remarks />
            public string Id { get; set; }

            /// <remarks />
            public string Name { get; set; }

            /// <remarks />
            public string Description { get; set; }

            /// <remarks />
            [XmlArrayItem("section", IsNullable = false)]
            public SuiteSection[] Sections { get; set; }
        }

        /// <remarks />
        [XmlType(AnonymousType = true)]
        public class SuiteSection
        {
            /// <remarks />
            public string Name { get; set; }

            /// <remarks />
            public string Description { get; set; }

            /// <remarks />
            [XmlArrayItem("case", IsNullable = false)]
            public SuiteSectionCase[] Cases { get; set; }

            /// <remarks />
            public SuiteSection[] Sections { get; set; }
        }

        /// <remarks />
        [XmlType(AnonymousType = true)]
        public class SuiteSectionCase
        {
            /// <remarks />
            public string Id { get; set; }

            /// <remarks />
            public string Title { get; set; }

            /// <remarks />
            public string Type { get; set; }

            /// <remarks />
            public string Priority { get; set; }

            /// <remarks />
            public string Estimate { get; set; }

            /// <remarks />
            public string References { get; set; }

            /// <remarks />
            public SuiteSectionCaseCustom Custom { get; set; }
        }

        /// <remarks />
        [XmlType(AnonymousType = true)]
        public class SuiteSectionCaseCustom
        {
            /// <remarks />
            public string Description { get; set; }

            /// <remarks />
            [XmlArrayItem("step", IsNullable = false)]
            public SuiteSectionCaseCustomStep[] StepsSeparated { get; set; }
        }

        /// <remarks />
        [XmlType(AnonymousType = true)]
        public class SuiteSectionCaseCustomStep
        {
            /// <remarks />
            public byte Index { get; set; }

            /// <remarks />
            public string Content { get; set; }

            /// <remarks />
            public string Expected { get; set; }
        }
    }
}