using System;
using System.Collections.Generic;
using System.Text;

namespace WordTableExtractor
{
    public class Tree<T>
    {
        public string Name { get; set; }
        public TreeNode<T> Root { get; set; }

        public bool HasRoot => Root != null;

        public Tree()
        {
            Root = new TreeNode<T>(default(T));
        }
    }


    public class TreeNode<T>
    {
        public TreeNode<T> Parent { get; set; }

        public List<TreeNode<T>> Children { get; set; } = new List<TreeNode<T>>();

        public bool IsRoot => Parent == null;

        public T Item { get; set; }

        public int Level
        {
            get
            {
                return GetLevel(this);
            }
        }

        public string Chapter
        {
            get
            {
                return GetChapter(this);
            }
        }

        public TreeNode<T> GetAnchestor(int level)
        {
            if (level == Level)
                return this;
            else
                return Parent.GetAnchestor(level);
        }

        public TreeNode(T item, TreeNode<T> parent = null)
        {
            Item = item;
            Parent = parent;
        }

        public TreeNode<T> AddChild(T item)
        {
            var childNode = new TreeNode<T>(item, this);
            Children.Add(childNode);

            return childNode;
        }

        public int GetLevel(TreeNode<T> node)
        {
            if (node.IsRoot)
                return 0;
            else
                return GetLevel(node.Parent) + 1;
        }

        public string GetChapter(TreeNode<T> node)
        {
            if(node.IsRoot)
            {
                return string.Empty;
            }
            else
            {
                var myIndex = node.Parent.Children.IndexOf(node);

                var item = node.Item as RequirementNode;
                var separator = item.IsHeading ? "." : "-";

                return GetChapter(node.Parent) + $"{separator}{myIndex}";
            }
        }

        public override string ToString()
        {
            return Item.ToString();
        }
    }
}
