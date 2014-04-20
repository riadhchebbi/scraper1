/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package com;

import java.util.List;

/**
 *
 * @author Riadh
 */
public class Category {
    
    private String s;
    private List<SubCategory> ls;

    public Category() {
    }

    public Category(String s, List<SubCategory> ls) {
        this.s = s;
        this.ls = ls;
    }

    public String getS() {
        return s;
    }

    public void setS(String s) {
        this.s = s;
    }

    public List<SubCategory> getLs() {
        return ls;
    }

    public void setLs(List<SubCategory> ls) {
        this.ls = ls;
    }

    @Override
    public String toString() {
        return s;
    }

   
    
    
}
