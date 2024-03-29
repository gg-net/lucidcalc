/* 
 * Copyright (C) 2014 GG-Net GmbH
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */
package eu.ggnet.lucidcalc;


import java.io.File;

import org.junit.jupiter.api.Test;

/**
 *
 */
public class SFormulaTest {

    @Test
    public void testFormulaError() {
        SFormula[] formulas =  {
          new SFormula(33,"+",22,"//"),
          new SFormula(33,"+",2,"/",0),
          new SFormula(33,"+",22,"/",Double.NaN),
          new SFormula(33,"+",22),
          new SFormula(33,"+",22),
        };
                        
        STable newTable = new STable();
        newTable.add(new STableColumn("TestFormula", 10));
        newTable.setModel(new STableModelList<>(formulas));

        SBlock block = new SBlock();
        block.add(new SFormula("SUMME(",newTable.getCellFirstRow(0),":",newTable.getCellLastRow(0),")"));
        
        
        CCalcDocument cdoc = new TempCalcDocument();
        cdoc.add(new CSheet("Sheet1",newTable,block));
        File f = LucidCalc.createWriter(LucidCalc.Backend.XLS).write(cdoc);
        f.delete();
//        System.out.println(f.toURI());
        // No assert needed. If Anything goes wrong in the write process, an exception would be thrown.
    }
        
}
