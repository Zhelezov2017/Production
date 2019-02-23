package ru.intertrust.cmj.af.so.misc;

import lotus.domino.Database;
import lotus.domino.Document;
import lotus.domino.NotesException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import ru.intertrust.cmj.af.core.AFCMDomino;
import ru.intertrust.cmj.af.core.AFSession;
import ru.intertrust.cmj.af.so.SOApplication;
import ru.intertrust.cmj.af.so.SOAppointment;
import ru.intertrust.cmj.af.so.SOBeard;
import ru.intertrust.cmj.af.utils.BeansUtils;
import ru.intertrust.cmj.dominodao.searcheval.af.so.SearchOrgsByNameAndInn;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

public class SPOExcelProcessorImpl implements SPOExcelProcessor{
    private File fileDouble = null;
    private Database db = null;
    private Workbook workBook = null;
    private long processCount = 0;
    private boolean ready = true;
    private List<Map<String,DuplicateReport>> tableEntry = new ArrayList<>();
    private Sheet she = null;

    public SPOExcelProcessorImpl(InputStream stream, String complect)  {
        try {
            workBook = WorkbookFactory.create(stream);
            db = AFCMDomino.getDbByIdent(AFCMDomino.AFDB_SYSTEM_ID_ORGDIRECTORY, complect);
            she = workBook.getSheetAt(0);
        } catch (NotesException | IOException | InvalidFormatException e) {
            throw new RuntimeException(e);
        }
    }


    public long getAllCount() {
        return she.getLastRowNum();
    }

    public long getProcessedCount() {
        return processCount;
    }


    public boolean isReady() {
        return ready;
    }

    public File getResult() {
        return fileDouble;
    }

    public void start(boolean doReplace) {
        ready = false;
        String cFullName;
        String cShortName;
        double innNumeric;
        String inn;
        String matching;
        Sheet sheet = workBook.getSheetAt(0);
        Iterator<Row> it = sheet.iterator();
        it.next();
        Map<String, SOBeard> person = new HashMap<String, SOBeard>();
        while (it.hasNext()) {
            Row row = it.next();
            List<SOBeard> requiredVizierBeards = new ArrayList<>();
            List<String> requiredVizierFIOs = new ArrayList<>();
            innNumeric = row.getCell(0).getNumericCellValue();
            cFullName = row.getCell(1).getStringCellValue();
            cShortName = row.getCell(2).getStringCellValue();
            if(row.getCell(3) == null){
                matching =  "";
            } else {
                matching =row.getCell(3).getStringCellValue();
            }
            inn = Integer.valueOf((int) innNumeric).toString();
            String[] listPerson = matching.replace(",", "").split("\\)");
            String[] listTubNum = matching.replace(",", "").split("\\)");
            for (int i = 0; i < listPerson.length - 1; i++) {
                listTubNum[i] = listTubNum[i].substring(listTubNum[i].lastIndexOf("("), listTubNum[i].lastIndexOf(""));
                listTubNum[i] = listTubNum[i].replace("(", "");
                listTubNum[i] = listTubNum[i].trim();
                listPerson[i] = listPerson[i].substring(0, listPerson[i].lastIndexOf("("));
                listPerson[i] = listPerson[i].replace("null", "");
                listPerson[i] = listPerson[i].trim();
                listPerson[i] = listPerson[i].substring(0, listPerson[i].indexOf(' '));
            }
            for (int i = 0; i < listTubNum.length; i++) {
                if (person.containsKey(listTubNum[i])) {
                    requiredVizierBeards.add(person.get(listTubNum[i]));
                    requiredVizierFIOs.add(listPerson[i]);
                } else {
                    SOApplication so = AFSession.get().getApplication(SOApplication.class);
                    SOApplication.BeardsSelection builder = so.createBeardsSelection();
                    builder.addBeardTypes(SOBeard.Type.SYS_HUMAN);
                    builder.addBeardTypes(SOBeard.Type.SYS_HUMAN_HEAD);
                    builder.setNameContains(listPerson[i]);
                    final List<SOBeard> result = builder.select(1, 999);
                    for (SOBeard sob: result){
                        SOBeard.DocflowData b = sob.originalData();
                        SOAppointment c = (SOAppointment) b.getParty();
                        if (c.isPrimary()) {
                            if (listTubNum[i].equals(sob.getTabNum())) {
                                person.put(sob.getTabNum(), sob);
                                requiredVizierBeards.add(sob);
                                requiredVizierFIOs.add(sob.toString());
                            }
                        }
                    }
                }
            }
            SearchOrgsByNameAndInn searchOrgs = BeansUtils.getBean("SearchOrgsByNameAndInn");
            Document dubl = searchOrgs.search(db, cShortName, cFullName, inn);
            Date dateCr = AFSession.get().currentUser().getStartWorkTime();
            try {
                if(dubl == null) {
                    List <String> ListSOBeard = new ArrayList<>();
                    Document doc = db.createDocument();
                    for (int i=0;i<requiredVizierBeards.size();i++){
                        ListSOBeard.add(requiredVizierBeards.get(i).toString(SOBeard.ToStringFormat.CMDOMINO_STD));
                    }
                    AFCMDomino.replaceItemValue(doc, "cFullName", cFullName, true);
                    AFCMDomino.replaceItemValue(doc, "cShortName", cShortName, true);
                    AFCMDomino.replaceItemValue(doc, "parentCFullName", cFullName, true);
                    AFCMDomino.replaceItemValue(doc, "parentCShortName", cShortName, true);
                    AFCMDomino.replaceItemValue(doc, "inn", inn, true);
                    AFCMDomino.replaceItemValue(doc, "requiredVizierBeards", ListSOBeard, true);
                    AFCMDomino.replaceItemValue(doc, "requiredVizierFIOs", requiredVizierFIOs, true);
                    AFCMDomino.replaceItemValue(doc, "form", "h1", true);
                    AFCMDomino.replaceItemValue(doc, "firmid", doc.getUniversalID(), true);
                    AFCMDomino.replaceItemValue(doc, "ids", doc.getUniversalID(), true);
                    AFCMDomino.replaceItemValue(doc, "parent", doc.getUniversalID(), true);
                    AFCMDomino.replaceItemValue(doc, "del", "No", true);
                    AFCMDomino.replaceItemValue(doc, "isVex", "No", true);
                    AFCMDomino.replaceItemValue(doc, "Union", "No", true);
                    AFCMDomino.replaceItemValue(doc, "zaj", "Not", true);
                    AFCMDomino.replaceItemValue(doc, "idOrg", AFCMDomino.getSODbReplicaID(), true);
                    AFCMDomino.replaceItemValue(doc, "structID", AFCMDomino.getSODbReplicaID(), true);
                    AFCMDomino.replaceItemValue(doc, "idAuthor", AFSession.get().currentUser().getBeard().toString(SOBeard.ToStringFormat.CMDOMINO_STD), true);
                    AFCMDomino.replaceItemValue(doc, "nativeNet", AFSession.get().getCurrentNet().name(), true);
                    AFCMDomino.replaceItemValue(doc, "corp_Serv", db.getServer(), true);
                    AFCMDomino.replaceItemValue(doc, "myServs", db.getServer(), true);
                    AFCMDomino.replaceItemValue(doc, "dateCr", dateCr, true);
                    AFCMDomino.replaceItemValue(doc, "date", dateCr, true);
                    doc.save();
                } else{
                    if (doReplace){
                        AFCMDomino.replaceItemValue(dubl, "cFullName", cFullName, true);
                        AFCMDomino.replaceItemValue(dubl, "cShortName", cShortName, true);
                        AFCMDomino.replaceItemValue(dubl, "parentCFullName", cFullName, true);
                        AFCMDomino.replaceItemValue(dubl, "parentCShortName", cShortName, true);
                        AFCMDomino.replaceItemValue(dubl, "inn", inn, true);
                        AFCMDomino.replaceItemValue(dubl, "requiredVizierBeards", requiredVizierBeards, true);
                        AFCMDomino.replaceItemValue(dubl, "requiredVizierFIOs", requiredVizierFIOs, true);
                        dubl.save();
                    } else {
                        Map<String, DuplicateReport> doubl = new HashMap<String, DuplicateReport>();
                        DuplicateReport innCopy = new DuplicateReport(inn,inn.equals(dubl.getItemValueString("inn")));
                        doubl.put("inn",innCopy);
                        DuplicateReport cShortNameCopy = new DuplicateReport(cShortName,cShortName.equals(dubl.getItemValueString("cShortName")));
                        doubl.put("cShortName",cShortNameCopy);
                        DuplicateReport cFullNameCopy = new DuplicateReport(cFullName,cFullName.equals(dubl.getItemValueString("cFullName")));
                        doubl.put("cFullName",cFullNameCopy);
                        DuplicateReport matchingCopy  = new DuplicateReport(matching, false);
                        doubl.put("matching",matchingCopy);
                        tableEntry.add(doubl);
                    }
                }
            } catch (NotesException e) {
                throw new RuntimeException(e);
            }
            processCount++;
        }
        try {
            fileDouble = File.createTempFile("excel_import_",".html");
            fileDubl(fileDouble);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        ready = true;
    }


    private void fileDubl(File fileToWrite){
        int tor = tableEntry.size();
        tor = 30 + tor *10;
        StringBuilder someWord =new StringBuilder( "<table border=\"1\" cellspacing=\"0\" cellpadding=\"15\" width=\"90%\" height=\""+tor+"\">\n" +
                "    <tr>\n" +
                "        <td bgcolor=\"#C0C0C0\"><b>ИНН</b></td>\n" +
                "        <td bgcolor=\"#C0C0C0\"><b>Полное наименование</b></td>\n" +
                "        <td bgcolor=\"#C0C0C0\"><b>Краткое наименование</b></td>\n" +
                "        <td bgcolor=\"#C0C0C0\"><b>Дополнительные согласующие</b></td>\n" +
                "    </tr>\n");
        try {
            FileWriter writer = new FileWriter(fileToWrite);
            for(int i=0;i<tableEntry.size();i++) {
                someWord.append("    <tr>\n");
                someWord.append(reportFunction(tableEntry.get(i).get("inn")));
                someWord.append(reportFunction(tableEntry.get(i).get("cFullName")));
                someWord.append(reportFunction(tableEntry.get(i).get("cShortName")));
                someWord.append(reportFunction(tableEntry.get(i).get("matching")));
                someWord.append("    </tr>\n");
            }
            someWord.append("</table>");
            String htmlCode = "<html><body><b>" + someWord + "</b></body></html>";
            writer.write(htmlCode);
            writer.flush();
            writer.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private String reportFunction(DuplicateReport obj){
        String someWord = "";
        if (obj.dubl) {
            someWord += "        <td  style=\"color:#ff0000\">" + obj.field + "</td>\n";
        } else {
            someWord += "        <td ><b>" + obj.field + "</b></td>\n";
        }
        return someWord;
    }

    private static class DuplicateReport {
        private String field;
        private Boolean dubl;
        public DuplicateReport(String field, Boolean dubl) {
            this.field = field;
            this.dubl = dubl;
        }
    }
}
