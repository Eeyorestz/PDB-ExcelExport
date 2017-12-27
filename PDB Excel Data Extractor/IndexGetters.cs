using System;
using System.Collections.Generic;
using System.Data;

namespace PDB_Excel_Data_Extractor
{
    public class IndexGetters
    {
        public  List<List<int>> Ranges(DataTable dataTable)
        {
            List<int> listOf = ListOfAllRanges(dataTable);
            List<List<int>> list = new List<List<int>>();
            for (int i = 0; i < listOf.Count-1; i++)
            {
                List<int> tempList = new List<int>();
                tempList.Add(listOf[i]);
                tempList.Add(listOf[i+1]);
                list.Add(tempList);
            }
            return list;
        }

        public List<string> listOfWorkouts(DataTable dataTable)
        {
            List<string> list = new List<string>();
            List<int> listOf = ListOfAllRanges(dataTable);
            for (int i = 0; i < listOf.Count; i++)
            {
                list.Add(dataTable.Rows[listOf[i] + 1][0].ToString());
            }
            return list;
        }

        private List<int> ListOfAllRanges(DataTable dataTable)
        {
            List<int> listOfIntegers = new List<int>();
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                if (dataTable.Rows[i][0].ToString().Contains("Кеш каса:"))
                {
                    listOfIntegers.Add(i - 2);
                }
            }
            return listOfIntegers;
        }

        public int RowOfCard(DataTable dataTable, string cardName)
        {
            int rowIndex = 0;
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                var ff = dataTable.Rows[i][0].ToString();
                if (dataTable.Rows[i][0].ToString().Equals(cardName))
                {
                    rowIndex = i;
                    break;
                }
            }
            return rowIndex;
        }

        public int CardExpirationStartingRowIndex(DataTable table)
        {
            int index = 0;
            for (int i = 0; i < table.Rows.Count; i++)
            {
                if (table.Rows[i][0].ToString().Contains("Номер на карта"))
                {
                    index = i;
                    break;
                }
            }
            return index;
        }
    }
}
