package com.gang.gselfservice.config;

import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

public class ThreadPool {

    private ThreadPool() {
    }

    private static class SingletonInstance {
        private static final ExecutorService INSTANCE = Executors.newFixedThreadPool(5);
    }

    public static ExecutorService getExecutorService() {
        return SingletonInstance.INSTANCE;
    }
}
