package org.wesol.helper;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;

public class LinkHelper {
    public static XSSFHyperlink createHyperlink(HyperlinkType type) {
        Hyperlink link = new Hyperlink() {
            @Override
            public int getFirstRow() {
                return 0;
            }

            @Override
            public void setFirstRow(int i) {

            }

            @Override
            public int getLastRow() {
                return 0;
            }

            @Override
            public void setLastRow(int i) {

            }

            @Override
            public int getFirstColumn() {
                return 0;
            }

            @Override
            public void setFirstColumn(int i) {

            }

            @Override
            public int getLastColumn() {
                return 0;
            }

            @Override
            public void setLastColumn(int i) {

            }

            @Override
            public String getAddress() {
                return null;
            }

            @Override
            public void setAddress(String s) {

            }

            @Override
            public String getLabel() {
                return null;
            }

            @Override
            public void setLabel(String s) {

            }

            @Override
            public HyperlinkType getType() {
                return type;
            }
        };
        return new XSSFHyperlink(link);
    }

}
