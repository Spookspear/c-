
        private void delLinesModeB(Excel.Worksheet Wks)
        {
            var range = Wks.UsedRange;
            range.SpecialCells(XlCellType.xlCellTypeConstants).EntireRow.Hidden = true;
            range.SpecialCells(XlCellType.xlCellTypeVisible).Delete(XlDeleteShiftDirection.xlShiftUp);
            range.EntireRow.Hidden = false;


            // then loop down from top
            
            for (int i = 1; i < intTotalRows; i++)
            {
            	if isEmptyCell((Wks.Cells[i,1].Value))
            }
            
            stopping at 1st blank - for the number
            create range to the end
            and delete it



        }
