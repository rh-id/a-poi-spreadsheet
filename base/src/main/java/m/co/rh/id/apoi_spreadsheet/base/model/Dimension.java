package m.co.rh.id.apoi_spreadsheet.base.model;

public class Dimension implements Cloneable {

    public int width;

    public int height;

    public Dimension() {
        this(0, 0);
    }

    public Dimension(int width, int height) {
        this.width = width;
        this.height = height;
    }

    public double getHeight() {
        return height;
    }

    public double getWidth() {
        return width;
    }

    public void setSize(Dimension dimension) {
        setSize(dimension.width, dimension.height);
    }

    public void setSize(int width, int height) {
        this.width = width;
        this.height = height;
    }

    public void setSize(double width, double height) {
        this.width = (int) width;
        this.height = (int) height;
    }

}
