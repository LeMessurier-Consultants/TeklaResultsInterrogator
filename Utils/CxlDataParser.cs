using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace TeklaResultsInterrogator.Utils
{
    public class CxlDataParser
    {
        public string FilePath { get; }

        public List<MemberData> Members { get; private set; } = new List<MemberData>();
        public Dictionary<int, double> NodeZCoordinates { get; private set; } = new Dictionary<int, double>();
        public List<LevelData> Levels { get; private set; } = new List<LevelData>();
        public List<LoadCaseData> LoadCases { get; private set; } = new List<LoadCaseData>();
        public List<BaseForceData> BaseForces { get; private set; } = new List<BaseForceData>();

        public CxlDataParser(string filePath)
        {
            FilePath = filePath ?? throw new ArgumentNullException(nameof(filePath));
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"CXL file not found: {filePath}");
        }

        public void Parse()
        {
            XDocument doc = XDocument.Load(FilePath);

            ParseNodes(doc);
            ParseLevels(doc);
            ParseMembers(doc);
            ParseLoadCases(doc);
            ParseBaseForces(doc);
        }

        private void ParseNodes(XDocument doc)
        {
            NodeZCoordinates.Clear();

            var coords = doc.Descendants("Coord");
            foreach (var c in coords)
            {
                if (int.TryParse(c.Attribute("Id")?.Value, out int nodeId))
                {
                    var zStr = c.Element("Z")?.Value;
                    if (double.TryParse(zStr, out double z))
                    {
                        NodeZCoordinates[nodeId] = z;
                    }
                }
            }
        }

        private void ParseLevels(XDocument doc)
        {
            Levels.Clear();

            var uniqueZs = NodeZCoordinates.Values.Select(z => Math.Round(z, 2)).Distinct().OrderBy(z => z).ToList();
            int idx = 0;
            foreach (var z in uniqueZs)
            {
                Levels.Add(new LevelData { Name = $"Level_{idx + 1}", Elevation = z });
                idx++;
            }
        }

        private void ParseMembers(XDocument doc)
        {
            Members.Clear();

            var members = doc.Descendants("Member");
            foreach (var m in members)
            {
                string id = m.Attribute("MemberId")?.Value ?? "";
                int startNode = int.Parse(m.Element("StartNode")?.Value ?? "0");
                int endNode = int.Parse(m.Element("EndNode")?.Value ?? "0");
                string partMark = m.Descendants("PART_MARK").FirstOrDefault()?.Value ?? "";
                string designGrp = m.Descendants("DESIGN_GRP").FirstOrDefault()?.Value ?? "";
                string sectionSize = m.Element("SectionSize")?.Value ?? "";
                string materialGrade = m.Element("MaterialGrade")?.Value ?? "";
                string memberType = m.Descendants("MBR_TYPE").FirstOrDefault()?.Value ?? "";

                Members.Add(new MemberData
                {
                    Id = id,
                    StartNode = startNode,
                    EndNode = endNode,
                    PartMark = partMark,
                    DesignGroup = designGrp,
                    SectionSize = sectionSize,
                    MaterialGrade = materialGrade,
                    MemberType = memberType
                });
            }
        }

        private void ParseLoadCases(XDocument doc)
        {
            LoadCases.Clear();

            var loadcases = doc.Descendants("Case");
            foreach (var c in loadcases)
            {
                if (int.TryParse(c.Attribute("CaseNo")?.Value, out int caseNo))
                {
                    string title = c.Element("Title")?.Value ?? "";
                    string type = c.Element("Type")?.Value ?? "";

                    LoadCases.Add(new LoadCaseData
                    {
                        CaseNumber = caseNo,
                        Title = title,
                        Type = type
                    });
                }
            }
        }

        private void ParseBaseForces(XDocument doc)
        {
            BaseForces.Clear();

            var bfBases = doc.Descendants("BFBase");
            foreach (var bfBase in bfBases)
            {
                int nodeNo = int.Parse(bfBase.Attribute("NodeNo")?.Value ?? "0");
                string mark = bfBase.Attribute("Mark")?.Value ?? "";

                var bfCases = bfBase.Descendants("BFCase");
                foreach (var bfCase in bfCases)
                {
                    int caseNo = int.Parse(bfCase.Attribute("CaseNo")?.Value ?? "0");
                    double.TryParse(bfCase.Attribute("Fv")?.Value, out double fv);
                    double.TryParse(bfCase.Attribute("Fmaj")?.Value, out double fmaj);
                    double.TryParse(bfCase.Attribute("Fmin")?.Value, out double fmin);
                    double.TryParse(bfCase.Attribute("Mmaj")?.Value, out double mmaj);
                    double.TryParse(bfCase.Attribute("Mmin")?.Value, out double mmin);

                    BaseForces.Add(new BaseForceData
                    {
                        NodeNo = nodeNo,
                        Mark = mark,
                        CaseNumber = caseNo,
                        ShearMajor = fmaj,
                        ShearMinor = fmin,
                        MomentMajor = mmaj,
                        MomentMinor = mmin,
                        AxialForce = fv,
                        Torsion = 0
                    });
                }
            }
        }
    }

    public class MemberData
    {
        public string Id { get; set; }
        public int StartNode { get; set; }
        public int EndNode { get; set; }
        public string PartMark { get; set; } = "";
        public string DesignGroup { get; set; } = "";
        public string SectionSize { get; set; } = "";
        public string MaterialGrade { get; set; } = "";
        public string MemberType { get; set; } = "";
    }

    public class LevelData
    {
        public string Name { get; set; } = "";
        public double Elevation { get; set; }
    }

    public class LoadCaseData
    {
        public int CaseNumber { get; set; }
        public string Title { get; set; } = "";
        public string Type { get; set; } = "";
    }

    public class BaseForceData
    {
        public int NodeNo { get; set; }
        public string Mark { get; set; } = "";
        public int CaseNumber { get; set; }
        public double ShearMajor { get; set; }
        public double ShearMinor { get; set; }
        public double MomentMajor { get; set; }
        public double MomentMinor { get; set; }
        public double AxialForce { get; set; }
        public double Torsion { get; set; }
    }
}
