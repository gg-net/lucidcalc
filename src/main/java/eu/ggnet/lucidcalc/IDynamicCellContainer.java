/*
 * Copyright (C) 2014 GG-Net GmbH - Oliver Günther.
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 3 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; If not, see <http://www.gnu.org/licenses/>.
 */
package eu.ggnet.lucidcalc;

import java.util.Set;

/**
 * Represents something that contains cells.
 */
public interface IDynamicCellContainer {

    /**
     * Returns all Cells of the Container, but shifted absolutely to row and column index
     *
     * @param columnIndex the column index
     * @param rowIndex    the row index
     * @return a set of Cells shifted.
     */
    Set<CCell> getCellsShiftedTo(int columnIndex, int rowIndex);

    CCellComposite shiftTo(int toColumnIndex, int toRowIndex);

    int getRowCount();

    int getColumnCount();

}
