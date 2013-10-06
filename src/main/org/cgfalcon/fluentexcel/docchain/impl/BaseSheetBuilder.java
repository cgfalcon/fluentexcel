package org.cgfalcon.fluentexcel.docchain.impl;

import org.cgfalcon.fluentexcel.docchain.BlockBuilder;
import org.cgfalcon.fluentexcel.docchain.DocBuilder;
import org.cgfalcon.fluentexcel.docchain.SheetBuilder;
import org.cgfalcon.fluentexcel.entity.*;

import java.util.Map;

/**
 * User: falcon.chu
 * Date: 13-9-10
 * Time: 下午3:40
 */
public class BaseSheetBuilder implements SheetBuilder {

    private DocBuilder docBuilder;
    private DataSheet dataSheet = new DataSheet();

    public BaseSheetBuilder(BaseDocBuilder styleDocBuilder) {
        this.docBuilder = styleDocBuilder;
    }

    @Override
    public SheetBuilder sheetName(String sheetName) {
        dataSheet.setSheetName(sheetName);
        return this;
    }

    @Override
    public SheetBuilder withColWidth(Map<Integer, Integer> colWidth) {
        dataSheet.setColWidth(colWidth);
        return this;
    }

    @Override
    public SheetBuilder addMergerRegion(MergerRegion mergerRegion) {
        dataSheet.addMergerRegion(mergerRegion);
        return this;
    }


    @Override
    public SheetBuilder addBlock(DataBlock block) {
        dataSheet.addBlock(block);
        return this;
    }

    @Override
    public BlockBuilder createBlock() {
        return new BaseBlockBuilder(this);
    }

    @Override
    public DocBuilder sheetOver() {
        docBuilder.addSheet(dataSheet);
        return docBuilder;  //To change body of implemented methods use File | Settings | File Templates.
    }

    @Override
    public int blockCount() {
        return dataSheet.getBlocks().size();
    }
}
