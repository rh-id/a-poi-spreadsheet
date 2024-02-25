package m.co.rh.id.apoi_spreadsheet;

import androidx.test.internal.runner.junit4.AndroidJUnit4ClassRunner;
import androidx.test.internal.util.AndroidRunnerParams;
import androidx.test.platform.app.InstrumentationRegistry;

import org.junit.runner.notification.RunNotifier;
import org.junit.runners.model.FrameworkMethod;
import org.junit.runners.model.InitializationError;
import org.junit.runners.model.Statement;

import m.co.rh.id.apoi_spreadsheet.base.POISpreadsheetContext;

public class POIJUnit4ClassRunner extends AndroidJUnit4ClassRunner {

    public POIJUnit4ClassRunner(Class<?> klass) throws InitializationError {
        super(klass);
    }

    public POIJUnit4ClassRunner(Class<?> klass, AndroidRunnerParams runnerParams) throws InitializationError {
        super(klass, runnerParams);
    }

    @Override
    protected Statement classBlock(RunNotifier notifier) {
        Statement parentStatement = super.classBlock(notifier);
        return new Statement() {
            @Override
            public void evaluate() throws Throwable {
                POISpreadsheetContext.getInstance().setAppContext(InstrumentationRegistry.getInstrumentation().getTargetContext());
                parentStatement.evaluate();
                POISpreadsheetContext.getInstance().setAppContext(null);
            }
        };
    }

    @Override
    protected Statement methodBlock(FrameworkMethod method) {
        return new POIThreadStatement(super.methodBlock(method));
    }
}
