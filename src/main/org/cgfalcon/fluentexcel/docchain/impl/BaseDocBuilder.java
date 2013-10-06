package org.cgfalcon.fluentexcel.docchain.impl;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Workbook;
import org.cgfalcon.fluentexcel.docchain.DocBuilder;
import org.cgfalcon.fluentexcel.docchain.SheetBuilder;
import org.cgfalcon.fluentexcel.render.Render;
import org.cgfalcon.fluentexcel.entity.DataDoc;
import org.cgfalcon.fluentexcel.entity.DataSheet;

/**
 * User: falcon.chu
 * Date: 13-9-10
 * Time: 下午3:32
 */
public class BaseDocBuilder implements DocBuilder {

    private static Logger logger = Logger.getLogger(BaseDocBuilder.class);

    private Render render;
    private DataDoc dataDoc = new DataDoc();
    private String saveDir = "/tmp";



    public BaseDocBuilder() {

    }




    @Override
    public DocBuilder withDocRender(Render render) {
        this.render = render;
        return this;
    }

    @Override
    public DocBuilder type(String docType) {

        dataDoc.setDocType(docType);
        return this;
    }

    @Override
    public DocBuilder docName(String docName) {
        dataDoc.setDocName(docName);
        return this;
    }

    @Override
    public DocBuilder addSheet(DataSheet sheet) {
        dataDoc.addSheet(sheet);
        return this;
    }

    @Override
    public DocBuilder saveTo(String saveDir) {
        if (saveDir == null) {
            logger.info("Hence <saveDir> is null, so excel will be generate to /tmp ");
        }
        dataDoc.setDocDir(saveDir);
        return this;  //To change body of implemented methods use File | Settings | File Templates.
    }

    @Override
    public SheetBuilder createSheet() {
        return new BaseSheetBuilder(this);
    }

    @Override
    public String rendDoc() {
        if (render == null) {
            throw new IllegalStateException("Null Render Object. Please verify render has been set correctly!");
        }
        render.render(dataDoc);
        return render.produceExcel(dataDoc);
    }

    @Override
    public int sheetCount() {
        return dataDoc.getSheets().size();
    }
}
