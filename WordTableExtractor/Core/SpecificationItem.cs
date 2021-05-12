using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordTableExtractor.Core
{
    public class SpecificationItem : ITreeNode
    {
        public string TypeFieldName { get; private set; }
        public string HierarchyFieldName { get; private set; }
        public string ContentFieldName { get; private set; }

        public string Section { get; private set; }
        public string Hierarchy
        {
            get
            {
                if (string.IsNullOrEmpty(HierarchyFieldName))
                    return string.Empty;

                if (!Values.ContainsKey(HierarchyFieldName))
                    return string.Empty;

                return Values[HierarchyFieldName];
            }
        }
        public ArtifactType Type
        {
            get
            {
                if(string.IsNullOrEmpty(TypeFieldName))
                    return ArtifactType.Unknown;

                if (!Values.ContainsKey(TypeFieldName))
                    return ArtifactType.Unknown;

                var typeString = Values[TypeFieldName];
                if (Enum.TryParse<ArtifactType>(typeString, out ArtifactType parsedType))
                {
                    return parsedType;
                }
                else
                {
                    if (typeString.ToLower().Contains("info"))
                        return ArtifactType.Information;
                    else
                        return ArtifactType.Unknown;
                }
            }
        }
        public string Content
        {
            get
            {
                if (string.IsNullOrEmpty(ContentFieldName))
                    return string.Empty;

                if (!Values.ContainsKey(ContentFieldName))
                    return string.Empty;

                return Values[ContentFieldName].Trim();
            }
        }
        public string SourceRange { get; private set; }
        public RequirementNodeAddress Address
        {
            get
            {
                if (string.IsNullOrEmpty(Hierarchy))
                    return null;
                else
                    return new RequirementNodeAddress(Hierarchy);
            }
        }

        public Dictionary<string, string> Values { get; set; } = new Dictionary<string, string>();

        public SpecificationItem(Dictionary<string, string> fields, string typeFieldname, string hierarchyFieldName, string contentFieldName, string sourceRange = "")
        {
            Values = fields;

            TypeFieldName = typeFieldname;
            HierarchyFieldName = hierarchyFieldName;
            ContentFieldName = contentFieldName;
            SourceRange = sourceRange;
        }

        public bool IsHeading
        {
            get
            {
                return Type == ArtifactType.Heading;
            }
        }

        public bool IsInformation
        {
            get
            {
                return Type == ArtifactType.Information;
            }
        }

        public bool IsRequirement
        {
            get
            {
                return Type == ArtifactType.Requirement;
            }
        }        

        public string Title
        {
            get
            {
                if(IsHeading)
                {
                    return $"{Address} {Content}";
                }
                else
                {
                    return $"REQ {Address}";
                }
            }
        }

        public string Summary
        {
            get
            {
                if (IsHeading)
                {
                    return Content;
                }
                else
                {
                    return $"REQ {Address}";
                }
            }
        }

        public bool IsLeaf
        {
            get
            {
                return Address.IsLeaf;
            }
        }

        public int Level
        {
            get
            {
                return Address.Level;
            }
        }

        public override string ToString()
        {
            return ToString(SpecificationStringFormat.Full);
        }

        public string ToString(SpecificationStringFormat format = SpecificationStringFormat.Short)
        {
            switch(format)
            {
                case SpecificationStringFormat.Short:
                    return Hierarchy;
                case SpecificationStringFormat.Full:
                    var maxLen = 50;
                    string content;

                    if (!string.IsNullOrEmpty(Content))
                        content = Content.Length > maxLen ? Content.Substring(0, maxLen) + "..." : Content;
                    else
                        content = "<empty>";

                    return $"{Type} {Address} {content}";
                default:
                    return Hierarchy;
            }
        }
    }

    public class RequirementNodeAddress
    {
        private string _section;

        public RequirementNodeAddress(string section)
        {
            _section = section;
        }

        public bool IsValid
        {
            get
            {
                return !string.IsNullOrEmpty(_section);
            }
        }

        public List<int> Path
        {
            get
            {
                if (!IsValid)
                    return new List<int>();

                if (_section.Contains("-")) // is leaf
                {
                    var chunks = _section.Split('-');
                    var nodes = chunks[0];
                    var leaf = chunks[1];

                    return nodes.Split('.').Select(x => Convert.ToInt32(x)).ToList();
                }
                else
                {
                    if (_section.Contains("."))
                    {
                        try
                        {
                            return _section.Split('.').Select(x => Convert.ToInt32(x)).ToList();
                        }
                        catch
                        {
                            return new List<int>();
                        }
                    }
                    else
                    {
                        if (int.TryParse(_section, out int result))
                            return new List<int> { result };
                        else
                            return new List<int>();
                    }
                }
            }
        }

        public string PathString
        {
            get
            {
                return string.Join('.', Path);
            }
        }

        public int? Number
        {
            get
            {
                if (!IsValid)
                    return null;

                if (_section.Contains("-")) // is leaf
                {
                    var chunks = _section.Split('-');
                    var leaf = chunks[1];
                    return Convert.ToInt32(leaf);
                }
                else
                {
                    return null;
                }
            }
        }

        public int Level
        {
            get
            {
                var path = Path;
                return path.Count;
            }
        }

        public bool IsLeaf
        {
            get
            {
                return Number.HasValue;
            }
        }

        public override string ToString()
        {
            if (IsLeaf)
                return PathString + $"-{Number}";
            else
                return PathString;
        }

        public RequirementNodeAddress ParentAddress
        {
            get
            {
                if (IsLeaf)
                {
                    return new RequirementNodeAddress(PathString);
                }
                else
                {
                    var parentPath = string.Join('.', Path.Take(Path.Count - 1));
                    return new RequirementNodeAddress(parentPath);
                }
            }
        }

        public RequirementNodeAddress PreviousAddress
        {
            get
            {
                if (IsLeaf)
                {
                    if (Number == null || Number.Value == 0)
                        return null;

                    return new RequirementNodeAddress(PathString + $"-{Number.Value - 1}");
                }
                else
                {
                    var chapterNumber = Path[^1];
                    if (chapterNumber == 0)
                        return null;

                    var path = Path.Take(Path.Count - 1).ToList();
                    path.Add(chapterNumber - 1);

                    var previousPath = string.Join('.', path);
                    return new RequirementNodeAddress(previousPath);
                }
            }
        }

        public RequirementNodeAddress NextAddress
        {
            get
            {
                if (IsLeaf)
                {
                    if (Number == null)
                        return null;

                    return new RequirementNodeAddress(PathString + $"-{Number.Value + 1}");
                }
                else
                {
                    var chapterNumber = Path[^1];

                    var path = Path.Take(Path.Count - 1).ToList();
                    path.Add(chapterNumber + 1);

                    var nextPath = string.Join('.', path);
                    return new RequirementNodeAddress(nextPath);
                }
            }
        }

        public override bool Equals(object obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }

            return ToString() == obj.ToString();
        }

        public static bool operator ==(RequirementNodeAddress a, RequirementNodeAddress b)
        {
            return a.Equals(b);
        }

        public static bool operator !=(RequirementNodeAddress a, RequirementNodeAddress b)
        {
            return !a.Equals(b);
        }
    }

    public enum ArtifactType
    {
        Heading,
        Information,
        Requirement,
        Unknown
    }

    public enum SpecificationStringFormat
    {
        Short,
        Full
    }
}
