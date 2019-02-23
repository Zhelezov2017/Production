package ru.intertrust.cmj.af.so.misc;

import java.io.File;
import java.io.InputStream;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.transaction.annotation.Transactional;

import lotus.domino.NotesException;
import ru.intertrust.cmj.af.core.AFCMDomino;
import ru.intertrust.cmj.af.misc.AFRunnable;

/**
 * Класс для запуска импорта содержимого excel-таблицы в СпО указанного комплекта
 * Запускает импорт в отдельном потоке
 *
 * @author LShershneva
 *
 */
public class SPOExcelImporter {

    private final static Logger log = LoggerFactory.getLogger(SPOExcelImporter.class);

    private static SPOExcelImporter importer;
    private static SPOExcelProcessor processor;
    private static Exception error;
    private static String spoReplicaId;


    public static SPOExcelImporter getInstance() {
        if (importer == null) {
            importer = new SPOExcelImporter();
        }
        return importer;
    }

    public long getAllCount() {
        if(processor!=null)
            return processor.getAllCount();
        return 0;
    }

    public long getProcessedCount() {
        if(processor!=null)
            return processor.getProcessedCount();
        return 0;
    }

    public File getResult() {
        if(processor!=null)
            return processor.getResult();
        return null;
    }

    public Exception getError() {
        return error;
    }

    public boolean isImportInProcess() {
        if(error!=null)
            return false;

        return (processor != null && !processor.isReady());
    }

    /**
     * replicaId обрабатываемой в данный момент БД СпО
     * @return replicaId СпО, если запущен процесс обработки, иначе null
     */
    public String getProcessedSpoRepId() {
        if(this.isImportInProcess()) {
            return spoReplicaId;
        }
        return null;
    }


    public void importFile(final InputStream excel, final String complect, final boolean doReplace) {
        log.debug("start import for {}, mode=", complect, doReplace?"replace":"add");
        if(this.isImportInProcess()) {
            throw(new RuntimeException("Процесс импорта уже запущен!"));
        }

        try {
            AFCMDomino.DbInfo dbInfo = AFCMDomino.getDbInfoByIdentNamed(AFCMDomino.AFDB_SYSTEM_ID_ORGDIRECTORY, complect);
            if(dbInfo == null || !complect.equals(dbInfo.complect)) {
                throw(new RuntimeException("Не удалось найти Справочник Организаций в комплекте "+complect));
            }
            spoReplicaId = dbInfo.replicaID;
        } catch (NotesException e) {
            throw(new RuntimeException(e));
        }

        processor = new SPOExcelProcessorDummy(excel, complect); //TODO - исправить на рабочий класс
        log.debug("processor added: "+processor);
        synchronized(processor) {
            error = null;

            //TODO synchronized?
            //TODO SPOExcelProcessor - на этапе отладки сделать Runnable или AFRunnable
            runTask(new Runnable() {
                @Override
                public void run() {
                    try {
                        processor.start(doReplace);
                        spoReplicaId = null;
                    }catch(Exception e) {
                        error = e;
                    }
                }
            });
        }
    }

    private static void runTask(final Runnable task) {
        new Thread(new AFRunnable(false) {
            @Transactional
            @Override
            protected void afRun() {
                task.run();
            }
        }).start();
    }


    /**
     * временный класс для проверок
     * @author LShershneva
     *
     */
    private static class SPOExcelProcessorDummy implements SPOExcelProcessor {

        private long allCount;
        private long processedCount;
        private boolean isReady;

        @Override
        public long getAllCount() {return allCount;}

        @Override
        public long getProcessedCount() {return processedCount;}

        @Override
        public boolean isReady() {return isReady;}

        @Override
        public File getResult() {return null;}

        public SPOExcelProcessorDummy(InputStream excel, String complect) {
            isReady = true;
            allCount = 30;
            processedCount = 0;
        }

        @Override
        public void start(boolean doReplace) {
            isReady = false;
            while(processedCount<allCount) {
                log.info("processing "+processedCount+ " out of "+allCount);
                processedCount++;
                try {
                    Thread.sleep(2000);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
            }
            isReady = true;
        }
    }

}
