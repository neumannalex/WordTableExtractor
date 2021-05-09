using System;
using System.Collections.Generic;
using System.Text;

namespace WordTableExtractor.Import
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

        public string GetDump()
        {
            var sb = new StringBuilder();

            sb.AppendLine("Tree");
            sb.AppendLine(Root.GetDump());

            return sb.ToString();
        }
    }


    public class TreeNode<T>
    {
        public TreeNode<T> Parent { get; set; }

        public List<TreeNode<T>> Children { get; set; } = new List<TreeNode<T>>();

        public bool IsRoot => Parent == null;

        public T Item { get; set; }

        public int Count
        {
            get
            {
                return GetCount(this);
            }
        }

        public int Level
        {
            get
            {
                return GetLevel(this);
            }
        }

        public TreeNode<T> GetAnchestor(int level)
        {
            if (level < 0)
                return null;

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

        private int GetLevel(TreeNode<T> node)
        {
            if (node.IsRoot)
                return 0;
            else
                return GetLevel(node.Parent) + 1;
        }

        private int GetCount(TreeNode<T> node)
        {
            var count = IsRoot ? 0 : 1;
            foreach (var child in node.Children)
            {
                count += GetCount(child);
            }
            return count;
        }
        
        public override string ToString()
        {
            return Item.ToString();
        }

        public string GetDump(string indent = "")
        {
            string _cross = " ├─";
            string _corner = " └─";
            string _vertical = " │ ";
            string _space = "   ";

            var sb = new StringBuilder();

            sb.Append(indent);

            if (Children.Count == 0)
            {
                sb.Append(_corner);
                indent += _space;
            }
            else
            {
                sb.Append(_cross);
                indent += _vertical;
            }

            if (Item != null)
                sb.AppendLine(Item.ToString());
            else
            {
                if (IsRoot)
                    sb.AppendLine("Root");
                else
                    sb.AppendLine($"Item for node at level {Level} is NULL.");
            }

            var numberOfChildren = Children.Count;
            for (var i = 0; i < numberOfChildren; i++)
            {
                var child = Children[i];
                sb.Append(child.GetDump(indent));
            }

            return sb.ToString();
        }

        public void Print(string indent = "")
        {
            var dump = GetDump(indent);

            Console.WriteLine(dump);
        }
    }
}
