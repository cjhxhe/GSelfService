package com.gang.gselfservice.task;

import com.gang.gselfservice.utils.DateUtils;
import com.gang.gselfservice.utils.FileUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.channels.FileChannel;
import java.util.concurrent.Callable;

public class SourceBackupTask implements Callable<Boolean> {

    private static final String BACKUP_FOLDER = "/backup";

    private String sourcePath;

    public SourceBackupTask(String sourcePath) {
        this.sourcePath = sourcePath;
    }

    @Override
    public Boolean call() {
        try {
            String userDirectory = System.getProperty("user.dir");
            File backupDir = new File(userDirectory + BACKUP_FOLDER);
            if (!backupDir.exists()) {
                backupDir.mkdir();
            }
            File sourceFile = new File(sourcePath);
            if (sourceFile.exists()) {
                String destPath = backupDir.getPath() + File.separator
                        + DateUtils.getCurrentDateTime(DateUtils.DATETIME_FORMATTER2)
                        + "." + FileUtils.getExtension(sourcePath);
                copyFile(sourceFile, new File(destPath));
            }
        } catch (Exception e) {
            // ignore
        }
        return Boolean.TRUE;
    }

    private static void copyFile(File source, File dest) throws IOException {
        FileChannel inputChannel = null;
        FileChannel outputChannel = null;
        try {
            inputChannel = new FileInputStream(source).getChannel();
            outputChannel = new FileOutputStream(dest).getChannel();
            outputChannel.transferFrom(inputChannel, 0, inputChannel.size());
        } finally {
            inputChannel.close();
            outputChannel.close();
        }
    }
}
