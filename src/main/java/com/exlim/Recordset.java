package com.exlim;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Recordset {
    private Map<Integer, String> headersMap = new HashMap<>();
    private List<Record> records = new ArrayList<>();

    Map<Integer, String> getHeadersMap() {
        return headersMap;
    }
    public void setRecord(Record records) {
        this.records.add(records);
    }
    public void setHeader(int key, String value) {
        headersMap.put(key, value);
    }

    public String getHeader(int key) {
        if (headersMap.containsKey(key))
            return headersMap.get(key);
        else {
            System.out.println("No key found : " + key);
        }
        return null;
    }

    public void printHeaders() {
        System.out.println(headersMap.toString());
    }

    public void printRecords() {
        for (Record record : this.records) {
            record.printRecord();
        }
    }

    public List<Record> getRecords() {
        return records;
    }

    public static class Record {
        private Map<String, String> row = new HashMap<>();
        public Map<String, String> getRow() {
            return row;
        }

        public void setKeyValue(String key, String value) {
            this.row.put(key, value);
        }
        public String getValue(String key) {
            return this.row.get(key.toUpperCase().trim());
        }
        public void printRecord() {
            System.out.println(getRow().entrySet().toString());
        }
    }
}
