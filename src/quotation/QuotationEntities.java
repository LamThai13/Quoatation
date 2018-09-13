/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package quotation;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

/**
 *
 * @author Kevin
 */
public class QuotationEntities {
    private String SN;
    private String RMA;
    private String Model;
    private String Box;
    private String foam;
    private String powercord;
    private String DVI;
    private String keyboard;
    private String videocable;
    private String top;
    private String bottom;
    private String Front;
    private String warranty;
    private String portCover;
    private String partNo;
    private boolean Salereturn;
    private String date;
    private String HDD;
    private String BootUpPass;
    private String BootUpFail;
    private String Conclusion;
    private String memo;
    private String photo;
    

    public QuotationEntities(String SN, String RMA, String Model, String Box, String foam, String powercord, String DVI, String keyboard, String videocable, String top, String bottom, String Front, String warranty, String portCover, String partNo, boolean Salereturn, String date, String HDD, String BootUpPass, String BootUpFail, String Conclusion, String memo, String photo) {
        this.SN = SN;
        this.RMA = RMA;
        this.Model = Model;
        this.Box = Box;
        this.foam = foam;
        this.powercord = powercord;
        this.DVI = DVI;
        this.keyboard = keyboard;
        this.videocable = videocable;
        this.top = top;
        this.bottom = bottom;
        this.Front = Front;
        this.warranty = warranty;
        this.portCover = portCover;
        this.partNo = partNo;
        this.Salereturn = Salereturn;
        this.date = date;
        this.HDD = HDD;
        this.BootUpPass = BootUpPass;
        this.BootUpFail = BootUpFail;
        this.Conclusion = Conclusion;
        this.memo = memo;
        this.photo = photo;
    }

    public QuotationEntities() {
    }

    public String getSN() {
        return SN;
    }

    public void setSN(String SN) {
        this.SN = SN;
    }

    public String getPortCover() {
        return portCover;
    }

    public void setPortCover(String portCover) {
        this.portCover = portCover;
    }

    public String getRMA() {
        return RMA;
    }

    public String getPartNo() {
        return partNo;
    }

    public void setPartNo(String partNo) {
        this.partNo = partNo;
    }

    public boolean isSalereturn() {
        return Salereturn;
    }

    public void setSalereturn(boolean Salereturn) {
        this.Salereturn = Salereturn;
    }

    public String getDate() {
        return date;
    }

    public void setDate(String date) {
        this.date = date;
    }

    public String getHDD() {
        return HDD;
    }

    public void setHDD(String HDD) {
        this.HDD = HDD;
    }

    public String getBootUpPass() {
        return BootUpPass;
    }

    public void setBootUpPass(String BootUpPass) {
        this.BootUpPass = BootUpPass;
    }

    public String getBootUpFail() {
        return BootUpFail;
    }

    public void setBootUpFail(String BootUpFail) {
        this.BootUpFail = BootUpFail;
    }

    public String getConclusion() {
        return Conclusion;
    }

    public void setConclusion(String Conclusion) {
        this.Conclusion = Conclusion;
    }

    public String getMemo() {
        return memo;
    }

    public void setMemo(String memo) {
        this.memo = memo;
    }

    public String getPhoto() {
        return photo;
    }

    public void setPhoto(String photo) {
        this.photo = photo;
    }

    public void setRMA(String RMA) {
        this.RMA = RMA;
    }

    public String getModel() {
        return Model;
    }

    public void setModel(String Model) {
        this.Model = Model;
    }

    public String getBox() {
        return Box;
    }

    public void setBox(String Box) {
        this.Box = Box;
    }

    public String getFoam() {
        return foam;
    }

    public void setFoam(String foam) {
        this.foam = foam;
    }

    public String getPowercord() {
        return powercord;
    }

    public void setPowercord(String powercord) {
        this.powercord = powercord;
    }

    public String getDVI() {
        return DVI;
    }

    public void setDVI(String DVI) {
        this.DVI = DVI;
    }

    public String getKeyboard() {
        return keyboard;
    }

    public void setKeyboard(String keyboard) {
        this.keyboard = keyboard;
    }

    public String getVideocable() {
        return videocable;
    }

    public void setVideocable(String videocable) {
        this.videocable = videocable;
    }

    public String getTop() {
        return top;
    }

    public void setTop(String top) {
        this.top = top;
    }

    public String getBottom() {
        return bottom;
    }

    public void setBottom(String bottom) {
        this.bottom = bottom;
    }

    public String getFront() {
        return Front;
    }

    public void setFront(String Front) {
        this.Front = Front;
    }

    public String getWarranty() {
        return warranty;
    }

    public void setWarranty(String warranty) {
        this.warranty = warranty;
    }

    
    
}
