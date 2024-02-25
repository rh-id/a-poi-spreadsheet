package m.co.rh.id.apoi_spreadsheet.base.model;

public abstract class Dimension2D implements Cloneable {
    protected Dimension2D() {
    }

    public abstract double getHeight();

    public abstract double getWidth();

    public void setSize(Dimension2D dimension2D) {
        setSize(dimension2D.getWidth(), dimension2D.getHeight());
    }

    public abstract void setSize(double width, double height);
}
