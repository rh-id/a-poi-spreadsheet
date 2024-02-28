package m.co.rh.id.apoi_spreadsheet;

import androidx.test.platform.app.InstrumentationRegistry;

import org.junit.runner.Runner;
import org.junit.runner.notification.RunNotifier;
import org.junit.runners.Parameterized;
import org.junit.runners.model.Statement;

import m.co.rh.id.apoi_spreadsheet.base.POISpreadsheetContext;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.util.ThreadLocalUtil;

public class POIJUnit4Parameterized extends Parameterized {

    /**
     * Hack class to set POI spreadsheet context before calling parameterized parameter
     */
    private static class ContextSetup {
        private ContextSetup() {
            POISpreadsheetContext.getInstance().setAppContext(InstrumentationRegistry.getInstrumentation().getTargetContext());
        }
    }

    /**
     * Only called reflectively. Do not use programmatically.
     *
     * @param klass
     */
    public POIJUnit4Parameterized(Class<?> klass) throws Throwable {
        this(klass, new ContextSetup());
    }

    /**
     * Hack constructor to setup POI Spreadsheet Context
     */
    public POIJUnit4Parameterized(Class<?> klass, Object extra) throws Throwable {
        super(klass);
    }


    @Override
    protected Statement classBlock(RunNotifier notifier) {
        Statement parentStatement = super.classBlock(notifier);
        return new Statement() {
            @Override
            public void evaluate() throws Throwable {
                parentStatement.evaluate();
                POISpreadsheetContext.getInstance().setAppContext(null);
            }
        };
    }

    @Override
    protected void runChild(Runner runner, RunNotifier notifier) {
        POISpreadsheetContext.getInstance().executeAndWait(() -> {
            super.runChild(runner, notifier);
        });
    }
}
