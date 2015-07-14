/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.rockink.presentationhtml;

/**
 *
 * @author nirmal
 */
public class ShapeProperties {
    
    
    double x,y,cx,cy;

    public ShapeProperties(double x, double y, double cx, double cy) {
        this.x = x;
        this.y = y;
        this.cx = cx;
        this.cy = cy;
    }
    
    
    

    public double getX() {
        return x;
    }

    public void setX(double x) {
        this.x = x;
    }

    public double getY() {
        return y;
    }

    public void setY(double y) {
        this.y = y;
    }

    public double getCx() {
        return cx;
    }

    public void setCx(double cx) {
        this.cx = cx;
    }

    public double getCy() {
        return cy;
    }

    public void setCy(double cy) {
        this.cy = cy;
    }

    @Override
    public String toString() {
        return "ShapeProperties{" + "x=" + x + ", y=" + y + ", cx=" + cx + ", cy=" + cy + '}';
    }
    
    

    
    
    
}
