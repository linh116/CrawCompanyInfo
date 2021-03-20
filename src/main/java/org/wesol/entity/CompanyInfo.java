package org.wesol.entity;

public class CompanyInfo {
    String msdn;
    String companyName;
    String province;
    String fileName;

    public CompanyInfo(String msdn, String companyName, String province) {
        this.msdn = msdn;
        this.companyName = companyName;
        this.province = province;
    }

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public String getMsdn() {
        return msdn;
    }

    public void setMsdn(String msdn) {
        this.msdn = msdn;
    }

    public String getCompanyName() {
        return companyName;
    }

    public void setCompanyName(String companyName) {
        this.companyName = companyName;
    }

    public String getProvince() {
        return province;
    }

    public void setProvince(String province) {
        this.province = province;
    }
}
