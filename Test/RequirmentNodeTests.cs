using FluentAssertions;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using WordTableExtractor;
using Xunit;

namespace Test
{
    public class RequirmentNodeTests
    {
        private RequirementNode headerNode  = new RequirementNode("1.15", "Heading", "1.15 Closed Loop Simulation (HiL / SiL)");
        private RequirementNode subheaderNode0 = new RequirementNode("1.15.0", "Heading", "1.15.0 Summary");
        private RequirementNode subheaderNode1 = new RequirementNode("1.15.1", "Heading", "1.15.1 Terminology");
        private RequirementNode leafNode = new RequirementNode("1.15.0-1", "Information", "This document describes procedures");

        [Fact]
        public void Test1()
        {
            headerNode.IsHeading.Should().BeTrue();
            headerNode.Type.Should().Be(ArtifactType.Heading);

            subheaderNode0.IsHeading.Should().BeTrue();
            subheaderNode0.Type.Should().Be(ArtifactType.Heading);

            subheaderNode1.IsHeading.Should().BeTrue();
            subheaderNode1.Type.Should().Be(ArtifactType.Heading);

            leafNode.IsInformation.Should().BeTrue();
            leafNode.Type.Should().Be(ArtifactType.Information);

            var nodes = new List<RequirementNode> { headerNode, subheaderNode0, subheaderNode1, leafNode };

            var parent = nodes.Where(x => x.Address == leafNode.Address.ParentAddress).FirstOrDefault();
        }

        [Fact]
        public void Test2()
        {
            var items = new List<int>{1, 2, 3, 4, 5 };
            var test = items[^1];
        }
    }
}
