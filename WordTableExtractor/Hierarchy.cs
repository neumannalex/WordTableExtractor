using System;
using System.Collections.Generic;
using System.Text;

namespace WordTableExtractor
{
    public class Hierarchy
    {
        private List<int> _hierarchy;

        public Hierarchy()
        {
            _hierarchy = new List<int> { 1 };
        }

        public int Levels
        {
            get
            {
                return _hierarchy.Count;
            }
        }

        public void NewLevel()
        {
            _hierarchy.Add(1);
        }

        public bool HasLevel(int level)
        {
            return level > 0 && level <= _hierarchy.Count;
        }

        public void IncrementLevel(int level)
        {
            if (!HasLevel(level))
                throw new IndexOutOfRangeException();

            _hierarchy[level - 1]++;
        }

        public override string ToString()
        {
            return string.Join('.', _hierarchy);
        }
    }
}
