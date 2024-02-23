package m.co.rh.id.a_poi_spreadsheet.base.image;

import android.graphics.Canvas;
import android.graphics.Color;
import android.graphics.Matrix;
import android.graphics.Paint;
import android.graphics.RectF;

import java.util.LinkedHashMap;
import java.util.Map;

// java.awt.Graphics2D;
public class Graphics2D {
    private Matrix matrix;
    private Map<Object, Object> context;
    private Paint paint;
    private Canvas canvas;
    private Color color;

    public Graphics2D() {
        matrix = new Matrix();
        canvas = new Canvas();
        paint = new Paint();
        context = new LinkedHashMap<>();
    }

    public void setColor(Color color) {
        this.color = color;
    }

    public void setRenderingHint(Object key, Object value) {
        context.put(key, value);
    }

    public Object getRenderingHint(Object key) {
        return context.get(key);
    }

    public Matrix getTransform() {
        return new Matrix(matrix);
    }

    public void setTransform(Matrix matrix) {
        this.matrix = matrix;
    }

    public void translate(int x, int y) {
        canvas.translate(x, y);
        matrix.setTranslate(x, y);
    }

    public void scale(int x, int y) {
        canvas.scale(x, y);
        matrix.setScale(x, y);
    }

    public void rotate(double theta, double x, double y) {
        canvas.rotate((float) theta, (float) x, (float) y);
        matrix.setRotate((float) theta, (float) x, (float) y);
    }

    public void fillRect(int x, int y, int width, int height) {
        canvas.drawRect(x, y, width - x, height - y, paint);
        matrix.mapRect(new RectF(x, y, width - x, height - y));
    }
}
