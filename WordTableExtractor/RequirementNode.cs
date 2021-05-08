using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordTableExtractor
{
    public class RequirementNode
    {
        public string Section { get; private set; }
        public ArtifactType Type { get; private set; }
        public string Content { get; private set; }
        public string SourceRange { get; private set; }
        public RequirementNodeAddress Address { get; private set; }

        public Dictionary<string, object> Values { get; set; } = new Dictionary<string, object>();

        public RequirementNode(string section, string type, string content, string sourceRange = "")
        {
            Section = section;
            Content = content;
            SourceRange = sourceRange;

            if(!string.IsNullOrEmpty(type))
            {
                if (Enum.TryParse<ArtifactType>(type, out ArtifactType parsedType))
                {
                    Type = parsedType;
                }
                else
                {
                    if (type.ToLower().Contains("info"))
                        Type = ArtifactType.Information;
                    else
                        Type = ArtifactType.Unknown;
                }
            }
            else
            {
                Type = ArtifactType.Unknown;
            }

            Address = new RequirementNodeAddress(section);
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
                if (string.IsNullOrEmpty(Section))
                    return string.Empty;

                if (IsLeaf)
                    return $"REQ {Address}";
                else
                    return Content.Replace(Section, "").Trim();
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
                            return new List<int>(result);
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
                return Path.Count;
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
}
