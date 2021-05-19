using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordTableExtractor.Core
{
    public enum SectionizedTreeNodeType
    {
        Folder,
        Leaf
    }

    public class SectionizedTreeNode<T> : NeumannAlex.Tree.TreeNode<T>
    {
        public SectionizedTreeNode(T value, SectionizedTreeNodeType type = SectionizedTreeNodeType.Leaf) : base(value)
        {
            Type = type;
        }

        public SectionizedTreeNode(SectionizedTreeNodeType type = SectionizedTreeNodeType.Leaf) : base()
        {
            Type = type;
        }

        public SectionizedTreeNodeType Type { get; set; }

        public override List<int> Path
        {
            get
            {
                return GetMyPath(this);

                static List<int> GetMyPath(NeumannAlex.Tree.ITreeNode<T> node)
                {
                    if (node.IsRoot)
                    {
                        if (node.IsTreeRoot)
                            return new List<int>();
                        else
                            return new List<int> { 1 };
                    }
                    else
                    {
                        var path = GetMyPath(node.Parent);

                        var typedNode = node as SectionizedTreeNode<T>;
                        if(typedNode != null)
                        {
                            var nodeIndex = node.Parent.Children.Where(n => (n as SectionizedTreeNode<T>).Type == typedNode.Type).ToList().IndexOf(node) + 1;

                            path.Add(nodeIndex);
                        }                        

                        return path;
                    }
                }
            }
        }

        public override string PathString
        {
            get
            {
                if (this.Type == SectionizedTreeNodeType.Folder)
                    return string.Join('.', Path);
                else
                    return string.Join('.', Path.GetRange(0, Path.Count - 1)) + "-" + Path[^1];
            }
        }
    }
}
