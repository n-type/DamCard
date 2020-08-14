using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace DamCard
{
    public class DamNoData
    {
        private List<DamNo> nonTargetList = new List<DamNo>();
        private List<DamNo> priorityTargetList = new List<DamNo>();
        private List<DamNo> damNoList = new List<DamNo>();

        public DamNoData()
        {
        }

        public DamNoData(string directoryPath)
        {
            ReadNonTargetList(Path.Combine(directoryPath, CommonConstants.DamNoNonTargetFileName));
            ReadPriorityTargetList(Path.Combine(directoryPath, CommonConstants.DamNoPriorityTargetFileName));
            ReadDamNoList(Path.Combine(directoryPath, CommonConstants.DamNoFileName));
        }

        public void ReadNonTargetList(string filePath)
        {
            StreamReader streamReader = new StreamReader(filePath);
            streamReader.ReadLine();
            while (!streamReader.EndOfStream)
            {
                string line = streamReader.ReadLine();
                string[] values = line.Split(',');
                DamNo damNo = new DamNo
                {
                    No = int.Parse(values[0]),
                    Name = values[1],
                    RiverSystem = values[2],
                    River = values[3],
                    PrefectureCode = int.Parse(values[4]),
                    PrefectureName = values[5]
                };
                nonTargetList.Add(damNo);
            }
            streamReader.Close();
        }

        public void ReadPriorityTargetList(string filePath)
        {
            StreamReader streamReader = new StreamReader(filePath);
            streamReader.ReadLine();
            while (!streamReader.EndOfStream)
            {
                string line = streamReader.ReadLine();
                string[] values = line.Split(',');
                DamNo damNo = new DamNo
                {
                    No = int.Parse(values[0]),
                    Name = values[1],
                    RiverSystem = values[2],
                    River = values[3],
                    PrefectureCode = int.Parse(values[4]),
                    PrefectureName = values[5]
                };
                priorityTargetList.Add(damNo);
            }
            streamReader.Close();
        }

        public void ReadDamNoList(string filePath)
        {
            StreamReader streamReader = new StreamReader(filePath);
            streamReader.ReadLine();
            while (!streamReader.EndOfStream)
            {
                string line = streamReader.ReadLine();
                string[] values = line.Split(',');
                DamNo damNo = new DamNo
                {
                    No = int.Parse(values[0]),
                    Name = values[1],
                    RiverSystem = values[2],
                    River = values[3],
                    PrefectureCode = int.Parse(values[4]),
                    PrefectureName = values[5]
                };
                damNoList.Add(damNo);
            }
            streamReader.Close();
        }

        public int Search(string damName, string riverSystem, string prefectureName)
        {
            // 検索用に改行と括弧書きを取り除く
            var name = Regex.Replace(damName.Replace("\n", ""), @"[\(（\[［].+[\)）\]］]", "").Trim();

            // 検索対象外
            // 堰・沼などダム便覧未掲載
            var nonTarget = nonTargetList
                .Where(x => x.Name == name && x.RiverSystem == riverSystem && x.PrefectureName == prefectureName)
                .ToList();
            if (nonTarget.Count() > 0)
            {
                return 0;
            }

            // 優先検索
            // ダム名相違、再開発ダムなど
            var priorityTarget = priorityTargetList
                .Where(x => x.Name == name && x.RiverSystem == riverSystem && x.PrefectureName == prefectureName)
                .ToList();
            if (priorityTarget.Count() > 0)
            {
                return priorityTarget[0].No;
            }

            // 通常検索
            var result = damNoList
                .Where(x => x.Name.StartsWith(name))
                .Where(x => x.RiverSystem == riverSystem || x.PrefectureName == prefectureName)
                .ToList();

            if (result.Count == 1) return result[0].No;
            if (result.Count > 1)
            {
                // 同一都道府県内で同名のダムは水系名が一致するデータを使用
                foreach (var data in result)
                {
                    if (data.RiverSystem == riverSystem) return data.No;
                }
            }

            return 0;
        }
    }
}
