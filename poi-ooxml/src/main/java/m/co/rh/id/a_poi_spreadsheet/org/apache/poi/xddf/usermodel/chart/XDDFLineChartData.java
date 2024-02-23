/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

package m.co.rh.id.a_poi_spreadsheet.org.apache.poi.xddf.usermodel.chart;

import java.util.List;
import java.util.Map;

import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.util.Beta;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.util.Internal;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.xddf.usermodel.XDDFShapeProperties;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;

@Beta
public class XDDFLineChartData extends XDDFChartData {
    private CTLineChart chart;

    @Internal
    protected XDDFLineChartData(
            XDDFChart parent,
            CTLineChart chart,
            Map<Long, XDDFChartAxis> categories,
            Map<Long, XDDFValueAxis> values) {
        super(parent);
        this.chart = chart;
        for (CTLineSer series : chart.getSerList()) {
            this.series.add(new Series(series, series.getCat(), series.getVal()));
        }
        defineAxes(categories, values);
    }

    private void defineAxes(Map<Long, XDDFChartAxis> categories, Map<Long, XDDFValueAxis> values) {
        if (chart.sizeOfAxIdArray() == 0) {
            for (Long id : categories.keySet()) {
                chart.addNewAxId().setVal(id);
            }
            for (Long id : values.keySet()) {
                chart.addNewAxId().setVal(id);
            }
        }
        defineAxes(chart.getAxIdArray(), categories, values);
    }

    @Internal
    @Override
    protected void removeCTSeries(int n) {
        chart.removeSer(n);
    }

    @Override
    public void setVaryColors(Boolean varyColors) {
        if (varyColors == null) {
            if (chart.isSetVaryColors()) {
                chart.unsetVaryColors();
            }
        } else {
            if (chart.isSetVaryColors()) {
                chart.getVaryColors().setVal(varyColors);
            } else {
                chart.addNewVaryColors().setVal(varyColors);
            }
        }
    }

    public Grouping getGrouping() {
        return Grouping.valueOf(chart.getGrouping().getVal());
    }

   public void setGrouping(Grouping grouping) {
      if (chart.getGrouping() != null) {
         chart.getGrouping().setVal(grouping.underlying);
      } else {
         chart.addNewGrouping().setVal(grouping.underlying);
      }
   }

    @Override
    public XDDFChartData.Series addSeries(XDDFDataSource<?> category,
            XDDFNumericalDataSource<? extends Number> values) {
        final long index = this.parent.incrementSeriesCount();
        final CTLineSer ctSer = this.chart.addNewSer();
        ctSer.addNewCat();
        ctSer.addNewVal();
        ctSer.addNewIdx().setVal(index);
        ctSer.addNewOrder().setVal(index);
        final Series added = new Series(ctSer, category, values);
        this.series.add(added);
        return added;
    }

    public class Series extends XDDFChartData.Series {
        private CTLineSer series;

        protected Series(CTLineSer series, XDDFDataSource<?> category,
                XDDFNumericalDataSource<? extends Number> values) {
            super(category, values);
            this.series = series;
        }

        protected Series(CTLineSer series, CTAxDataSource category, CTNumDataSource values) {
            super(XDDFDataSourcesFactory.fromDataSource(category), XDDFDataSourcesFactory.fromDataSource(values));
            this.series = series;
        }

        /**
         * @since POI 5.2.3
         */
        public CTLineSer getCTLineSer() {
            return series;
        }

        @Override
        protected CTSerTx getSeriesText() {
            if (series.isSetTx()) {
                return series.getTx();
            } else {
                return series.addNewTx();
            }
        }

        @Override
        public void setShowLeaderLines(boolean showLeaderLines) {
            if (!series.isSetDLbls()) {
                series.addNewDLbls();
            }
            if (series.getDLbls().isSetShowLeaderLines()) {
                series.getDLbls().getShowLeaderLines().setVal(showLeaderLines);
            } else {
                series.getDLbls().addNewShowLeaderLines().setVal(showLeaderLines);
            }
        }

        @Override
        public XDDFShapeProperties getShapeProperties() {
            if (series.isSetSpPr()) {
                return new XDDFShapeProperties(series.getSpPr());
            } else {
                return null;
            }
        }

        @Override
        public void setShapeProperties(XDDFShapeProperties properties) {
            if (properties == null) {
                if (series.isSetSpPr()) {
                    series.unsetSpPr();
                }
            } else {
                if (series.isSetSpPr()) {
                    series.setSpPr(properties.getXmlObject());
                } else {
                    series.addNewSpPr().set(properties.getXmlObject());
                }
            }
        }

        /**
         * @since 4.0.1
         */
        public Boolean isSmooth() {
            if (series.isSetSmooth()) {
                return series.getSmooth().getVal();
            } else {
                return false;
            }
        }

        /**
         * @param smooth
         *        whether or not to smooth lines, if <code>null</code> then reverts to default.
         * @since 4.0.1
         */
        public void setSmooth(Boolean smooth) {
            if (smooth == null) {
                if (series.isSetSmooth()) {
                    series.unsetSmooth();
                }
            } else {
                if (series.isSetSmooth()) {
                    series.getSmooth().setVal(smooth);
                } else {
                    series.addNewSmooth().setVal(smooth);
                }
            }
        }

        /**
         * @param size
         * <dl><dt>Minimum inclusive:</dt><dd>2</dd><dt>Maximum inclusive:</dt><dd>72</dd></dl>
         */
        public void setMarkerSize(short size) {
            if (size < 2 || 72 < size) {
                throw new IllegalArgumentException("Minimum inclusive: 2; Maximum inclusive: 72");
            }
            CTMarker marker = getMarker();
            if (marker.isSetSize()) {
                marker.getSize().setVal(size);
            } else {
                marker.addNewSize().setVal(size);
            }
        }

        public void setMarkerStyle(MarkerStyle style) {
            CTMarker marker = getMarker();
            if (marker.isSetSymbol()) {
                marker.getSymbol().setVal(style.underlying);
            } else {
                marker.addNewSymbol().setVal(style.underlying);
            }
        }

        private CTMarker getMarker() {
            if (series.isSetMarker()) {
                return series.getMarker();
            } else {
                return series.addNewMarker();
            }
        }

        /**
         * @since 4.1.2
         */
        public boolean hasErrorBars() {
            return series.isSetErrBars();
        }

        /**
         * @since 4.1.2
         */
        public XDDFErrorBars getErrorBars() {
                if (series.isSetErrBars()) {
                    return new XDDFErrorBars(series.getErrBars());
                } else {
                    return null;
                }
        }

        /**
         * @since 4.1.2
         */
        public void setErrorBars(XDDFErrorBars bars) {
            if (bars == null) {
                if (series.isSetErrBars()) {
                    series.unsetErrBars();
                }
            } else {
                if (series.isSetErrBars()) {
                    series.getErrBars().set(bars.getXmlObject());
                } else {
                    series.addNewErrBars().set(bars.getXmlObject());
                }
            }
        }

        @Override
        protected CTAxDataSource getAxDS() {
            return series.getCat();
        }

        @Override
        protected CTNumDataSource getNumDS() {
            return series.getVal();
        }

        @Override
        protected void setIndex(long val) {
            series.getIdx().setVal(val);
        }

        @Override
        protected void setOrder(long val) {
            series.getOrder().setVal(val);
        }

        @Override
        protected List<CTDPt> getDPtList() {
            return series.getDPtList();
        }
    }
}
