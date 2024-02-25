package m.co.rh.id.apoi_spreadsheet;

import org.junit.runners.model.Statement;

import java.util.concurrent.atomic.AtomicReference;

import m.co.rh.id.apoi_spreadsheet.base.POISpreadsheetContext;

public class POIThreadStatement extends Statement {

    private Statement base;

    public POIThreadStatement(Statement base) {
        this.base = base;
    }

    @Override
    public void evaluate() throws Throwable {
        final AtomicReference<Throwable> exceptionRef = new AtomicReference<>();
        POISpreadsheetContext.getInstance().executeAndWait(() -> {
            try {
                base.evaluate();
            } catch (Throwable e) {
                exceptionRef.set(e);
            }
        });
        Throwable throwable = exceptionRef.get();
        if (throwable != null) {
            throw throwable;
        }
    }
}
