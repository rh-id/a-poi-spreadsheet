package m.co.rh.id.a_poi_spreadsheet.base;

import android.content.Context;
import android.util.Log;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.Callable;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Future;
import java.util.concurrent.FutureTask;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.TimeoutException;

/**
 * Context and ExecutorService to execute Spreadsheet related manipulation.
 * NOTE: This is a reused single thread executor, it may cause starvation deadlock just as other single thread executors.
 * DO NOT chain multiple futures that depends on each other to finish
 */
public class POISpreadsheetContext implements ExecutorService {
    private static POISpreadsheetContext instance;

    public static POISpreadsheetContext getInstance() {
        if (instance == null) {
            instance = new POISpreadsheetContext();
        }
        return instance;
    }

    private Context appContext;

    private final ReuseThread reuseThread;

    private POISpreadsheetContext() {
        reuseThread = new ReuseThread();
        // Android can be shutdown or killed so set this as daemon
        reuseThread.setDaemon(true);
    }

    public Context getAppContext() {
        return appContext;
    }

    public void setAppContext(Context context) {
        if (context != null) {
            appContext = context.getApplicationContext();
        } else {
            appContext = null;
        }
    }

    /**
     * No operation, this is a singleton unable to shutdown
     */
    @Override
    public void shutdown() {
        // Leave blank singleton unable to shutdown
    }

    /**
     * No operation, this is a singleton unable to shutdown
     */
    @Override
    public List<Runnable> shutdownNow() {
        // Leave blank singleton unable to shutdown
        return new ArrayList<>();
    }

    /**
     * No operation, this is a singleton unable to shutdown
     */
    @Override
    public boolean isShutdown() {
        // Leave blank singleton unable to shutdown
        return false;
    }

    /**
     * No operation, this is a singleton unable to shutdown
     */
    @Override
    public boolean isTerminated() {
        // Leave blank singleton unable to shutdown
        return false;
    }

    /**
     * No operation, this is a singleton unable to shutdown
     */
    @Override
    public boolean awaitTermination(long l, TimeUnit timeUnit) throws InterruptedException {
        // Leave blank singleton unable to shutdown
        return false;
    }

    public <T> Future<T> submit(Callable<T> callable) {
        checkStart();
        return reuseThread.submit(callable);
    }

    @Override
    public <T> Future<T> submit(Runnable runnable, T t) {
        checkStart();
        return reuseThread.submit(runnable, t);
    }

    @Override
    public Future<?> submit(Runnable runnable) {
        checkStart();
        return reuseThread.submit(runnable, null);
    }

    /**
     * Submit callable and wait for the execution
     */
    public <T> Future<T> submitAndWait(Callable<T> callable) {
        checkStart();
        return reuseThread.submitAndWait(callable);
    }

    @Override
    public <T> List<Future<T>> invokeAll(Collection<? extends Callable<T>> collection) {
        checkStart();
        List<Future<T>> futures = new ArrayList<>();
        for (Callable<T> callable : collection) {
            futures.add(reuseThread.submitAndWait(callable));
        }
        return futures;
    }

    @Override
    public <T> List<Future<T>> invokeAll(Collection<? extends Callable<T>> collection, long l, TimeUnit timeUnit) {
        checkStart();
        List<Future<T>> futures = new ArrayList<>();
        long startTimeMilis = System.currentTimeMillis();
        long endTimeMilis = startTimeMilis + TimeUnit.MILLISECONDS.convert(l, timeUnit);
        for (Callable<T> callable : collection) {
            long currentTimeMilis = System.currentTimeMillis();
            if (currentTimeMilis >= endTimeMilis) {
                FutureTask<T> futureTask = new FutureTask<>(callable);
                futureTask.cancel(true);
                futures.add(futureTask);
            } else {
                futures.add(reuseThread.submitAndWait(callable));
            }
        }
        return futures;

    }

    @Override
    public <T> T invokeAny(Collection<? extends Callable<T>> collection) throws ExecutionException, InterruptedException {
        checkStart();
        List<Future<T>> futures = new ArrayList<>();
        for (Callable<T> callable : collection) {
            futures.add(reuseThread.submit(callable));
        }
        T result = null;
        boolean findOne = false;
        try {
            while (!findOne) {
                for (Future<T> future : futures) {
                    if (future.isDone()) {
                        result = future.get();
                        findOne = true;
                        break;
                    }
                }
            }
        } finally {
            for (Future<T> future : futures) {
                future.cancel(true);
            }
        }
        return result;
    }

    @Override
    public <T> T invokeAny(Collection<? extends Callable<T>> collection, long l, TimeUnit timeUnit) throws ExecutionException, InterruptedException, TimeoutException {
        checkStart();
        List<Future<T>> futures = new ArrayList<>();
        for (Callable<T> callable : collection) {
            futures.add(reuseThread.submit(callable));
        }
        long startTimeMilis = System.currentTimeMillis();
        long endTimeMilis = startTimeMilis + TimeUnit.MILLISECONDS.convert(l, timeUnit);
        T result = null;
        boolean findOne = false;
        try {
            while (!findOne) {
                long currentTimeMilis = System.currentTimeMillis();
                if (currentTimeMilis >= endTimeMilis) {
                    throw new TimeoutException();
                }
                for (Future<T> future : futures) {
                    if (future.isDone()) {
                        result = future.get();
                        findOne = true;
                        break;
                    }
                }
            }
        } finally {
            for (Future<T> future : futures) {
                future.cancel(true);
            }
        }
        return result;
    }

    private void checkStart() {
        if (!reuseThread.isAlive()) {
            reuseThread.start();
        }
    }

    @Override
    public void execute(Runnable runnable) {
        checkStart();
        reuseThread.execute(runnable);
    }

    /**
     * Execute runnable and wait till finish
     */
    public void executeAndWait(Runnable runnable) {
        checkStart();
        reuseThread.executeAndWait(runnable);
    }


    private static class ReuseThread extends Thread {
        private BlockingQueue<FutureTask<?>> callableQueue;

        public ReuseThread() {
            super("poi-spreadsheet-thread");
        }

        @Override
        public synchronized void start() {
            if (callableQueue == null) {
                callableQueue = new ArrayBlockingQueue<>(Integer.MAX_VALUE);
            }
            super.start();
        }

        public <T> Future<T> submit(Callable<T> callable) {
            FutureTask<T> task = new FutureTask<>(callable);
            callableQueue.add(task);
            return task;
        }

        public <T> Future<T> submitAndWait(Callable<T> callable) {
            FutureTask<T> task = new FutureTask<>(callable);
            callableQueue.add(task);
            try {
                task.get();
            } catch (ExecutionException | InterruptedException e) {
                Log.e(getName(), e.getMessage(), e);
            }
            return task;
        }

        public <T> Future<T> submit(Runnable runnable, T t) {
            FutureTask<T> task = new FutureTask<>(runnable, t);
            callableQueue.add(task);
            return task;
        }

        public void execute(Runnable runnable) {
            FutureTask<Object> task = new FutureTask<>(runnable, null);
            callableQueue.add(task);
        }

        public void executeAndWait(Runnable runnable) {
            FutureTask<Object> task = new FutureTask<>(runnable, null);
            callableQueue.add(task);
            try {
                task.get();
            } catch (ExecutionException | InterruptedException e) {
                Log.e(getName(), e.getMessage(), e);
                throw new RuntimeException(e);
            }
        }

        @Override
        public void run() {
            // Thread is always alive as long as it hasn't finish this run
            // So this is intentional infinite loop
            while (isAlive()) {
                try {
                    FutureTask<?> task = callableQueue.take();
                    task.run();
                } catch (InterruptedException e) {
                    Log.e(getName(), e.getMessage(), e);
                }
            }
        }
    }
}
