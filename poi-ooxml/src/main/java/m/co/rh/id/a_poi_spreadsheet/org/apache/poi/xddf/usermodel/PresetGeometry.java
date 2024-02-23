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

package m.co.rh.id.a_poi_spreadsheet.org.apache.poi.xddf.usermodel;

import java.util.HashMap;

import org.openxmlformats.schemas.drawingml.x2006.main.STShapeType;

public enum PresetGeometry {
    ACCENT_BORDER_CALLOUT_1(STShapeType.ACCENT_BORDER_CALLOUT_1),
    ACCENT_BORDER_CALLOUT_2(STShapeType.ACCENT_BORDER_CALLOUT_2),
    ACCENT_BORDER_CALLOUT_3(STShapeType.ACCENT_BORDER_CALLOUT_3),
    ACCENT_CALLOUT_1(STShapeType.ACCENT_CALLOUT_1),
    ACCENT_CALLOUT_2(STShapeType.ACCENT_CALLOUT_2),
    ACCENT_CALLOUT_3(STShapeType.ACCENT_CALLOUT_3),
    ACTION_BUTTON_BACK_PREVIOUS(STShapeType.ACTION_BUTTON_BACK_PREVIOUS),
    ACTION_BUTTON_BEGINNING(STShapeType.ACTION_BUTTON_BEGINNING),
    ACTION_BUTTON_BLANK(STShapeType.ACTION_BUTTON_BLANK),
    ACTION_BUTTON_DOCUMENT(STShapeType.ACTION_BUTTON_DOCUMENT),
    ACTION_BUTTON_END(STShapeType.ACTION_BUTTON_END),
    ACTION_BUTTON_FORWARD_NEXT(STShapeType.ACTION_BUTTON_FORWARD_NEXT),
    ACTION_BUTTON_HELP(STShapeType.ACTION_BUTTON_HELP),
    ACTION_BUTTON_HOME(STShapeType.ACTION_BUTTON_HOME),
    ACTION_BUTTON_INFORMATION(STShapeType.ACTION_BUTTON_INFORMATION),
    ACTION_BUTTON_MOVIE(STShapeType.ACTION_BUTTON_MOVIE),
    ACTION_BUTTON_RETURN(STShapeType.ACTION_BUTTON_RETURN),
    ACTION_BUTTON_SOUND(STShapeType.ACTION_BUTTON_SOUND),
    ARC(STShapeType.ARC),
    BENT_ARROW(STShapeType.BENT_ARROW),
    BENT_CONNECTOR_2(STShapeType.BENT_CONNECTOR_2),
    BENT_CONNECTOR_3(STShapeType.BENT_CONNECTOR_3),
    DARK_BLUE(STShapeType.BENT_CONNECTOR_4),
    BENT_CONNECTOR_4(STShapeType.BENT_CONNECTOR_4),
    BENT_CONNECTOR_5(STShapeType.BENT_CONNECTOR_5),
    BENT_UP_ARROW(STShapeType.BENT_UP_ARROW),
    BEVEL(STShapeType.BEVEL),
    BLOCK_ARC(STShapeType.BLOCK_ARC),
    BORDER_CALLOUT_1(STShapeType.BORDER_CALLOUT_1),
    BORDER_CALLOUT_2(STShapeType.BORDER_CALLOUT_2),
    BORDER_CALLOUT_3(STShapeType.BORDER_CALLOUT_3),
    BRACE_PAIR(STShapeType.BRACE_PAIR),
    BRACKET_PAIR(STShapeType.BRACKET_PAIR),
    CALLOUT_1(STShapeType.CALLOUT_1),
    CALLOUT_2(STShapeType.CALLOUT_2),
    CALLOUT_3(STShapeType.CALLOUT_3),
    CAN(STShapeType.CAN),
    CHART_PLUS(STShapeType.CHART_PLUS),
    CHART_STAR(STShapeType.CHART_STAR),
    CHART_X(STShapeType.CHART_X),
    CHEVRON(STShapeType.CHEVRON),
    CHORD(STShapeType.CHORD),
    CIRCULAR_ARROW(STShapeType.CIRCULAR_ARROW),
    CLOUD(STShapeType.CLOUD),
    CLOUD_CALLOUT(STShapeType.CLOUD_CALLOUT),
    CORNER(STShapeType.CORNER),
    CORNER_TABS(STShapeType.CORNER_TABS),
    CUBE(STShapeType.CUBE),
    CURVED_CONNECTOR_2(STShapeType.CURVED_CONNECTOR_2),
    CURVED_CONNECTOR_3(STShapeType.CURVED_CONNECTOR_3),
    CURVED_CONNECTOR_4(STShapeType.CURVED_CONNECTOR_4),
    CURVED_CONNECTOR_5(STShapeType.CURVED_CONNECTOR_5),
    CURVED_DOWN_ARROW(STShapeType.CURVED_DOWN_ARROW),
    CURVED_LEFT_ARROW(STShapeType.CURVED_LEFT_ARROW),
    CURVED_RIGHT_ARROW(STShapeType.CURVED_RIGHT_ARROW),
    CURVED_UP_ARROW(STShapeType.CURVED_UP_ARROW),
    DECAGON(STShapeType.DECAGON),
    DIAGONAL_STRIPE(STShapeType.DIAG_STRIPE),
    DIAMOND(STShapeType.DIAMOND),
    DODECAGON(STShapeType.DODECAGON),
    DONUT(STShapeType.DONUT),
    DOUBLE_WAVE(STShapeType.DOUBLE_WAVE),
    DOWN_ARROW(STShapeType.DOWN_ARROW),
    DOWN_ARROW_CALLOUT(STShapeType.DOWN_ARROW_CALLOUT),
    ELLIPSE(STShapeType.ELLIPSE),
    ELLIPSE_RIBBON(STShapeType.ELLIPSE_RIBBON),
    ELLIPSE_RIBBON_2(STShapeType.ELLIPSE_RIBBON_2),
    FLOW_CHART_ALTERNATE_PROCESS(STShapeType.FLOW_CHART_ALTERNATE_PROCESS),
    FLOW_CHART_COLLATE(STShapeType.FLOW_CHART_COLLATE),
    FLOW_CHART_CONNECTOR(STShapeType.FLOW_CHART_CONNECTOR),
    FLOW_CHART_DECISION(STShapeType.FLOW_CHART_DECISION),
    FLOW_CHART_DELAY(STShapeType.FLOW_CHART_DELAY),
    FLOW_CHART_DISPLAY(STShapeType.FLOW_CHART_DISPLAY),
    FLOW_CHART_DOCUMENT(STShapeType.FLOW_CHART_DOCUMENT),
    FLOW_CHART_EXTRACT(STShapeType.FLOW_CHART_EXTRACT),
    FLOW_CHART_INPUT_OUTPUT(STShapeType.FLOW_CHART_INPUT_OUTPUT),
    FLOW_CHART_INTERNAL_STORAGE(STShapeType.FLOW_CHART_INTERNAL_STORAGE),
    FLOW_CHART_MAGNETIC_DISK(STShapeType.FLOW_CHART_MAGNETIC_DISK),
    FLOW_CHART_MAGNETIC_DRUM(STShapeType.FLOW_CHART_MAGNETIC_DRUM),
    FLOW_CHART_MAGNETIC_TAPE(STShapeType.FLOW_CHART_MAGNETIC_TAPE),
    FLOW_CHART_MANUAL_INPUT(STShapeType.FLOW_CHART_MANUAL_INPUT),
    FLOW_CHART_MANUAL_OPERATION(STShapeType.FLOW_CHART_MANUAL_OPERATION),
    FLOW_CHART_MERGE(STShapeType.FLOW_CHART_MERGE),
    FLOW_CHART_MULTIDOCUMENT(STShapeType.FLOW_CHART_MULTIDOCUMENT),
    FLOW_CHART_OFFLINE_STORAGE(STShapeType.FLOW_CHART_OFFLINE_STORAGE),
    FLOW_CHART_OFFPAGE_CONNECTOR(STShapeType.FLOW_CHART_OFFPAGE_CONNECTOR),
    FLOW_CHART_ONLINE_STORAGE(STShapeType.FLOW_CHART_ONLINE_STORAGE),
    FLOW_CHART_OR(STShapeType.FLOW_CHART_OR),
    FLOW_CHART_PREDEFINED_PROCESS(STShapeType.FLOW_CHART_PREDEFINED_PROCESS),
    FLOW_CHART_PREPARATION(STShapeType.FLOW_CHART_PREPARATION),
    FLOW_CHART_PROCESS(STShapeType.FLOW_CHART_PROCESS),
    FLOW_CHART_PUNCHED_CARD(STShapeType.FLOW_CHART_PUNCHED_CARD),
    FLOW_CHART_PUNCHED_TAPE(STShapeType.FLOW_CHART_PUNCHED_TAPE),
    FLOW_CHART_SORT(STShapeType.FLOW_CHART_SORT),
    FLOW_CHART_SUMMING_JUNCTION(STShapeType.FLOW_CHART_SUMMING_JUNCTION),
    FLOW_CHART_TERMINATOR(STShapeType.FLOW_CHART_TERMINATOR),
    FOLDED_CORNER(STShapeType.FOLDED_CORNER),
    FRAME(STShapeType.FRAME),
    FUNNEL(STShapeType.FUNNEL),
    GEAR_6(STShapeType.GEAR_6),
    GEAR_9(STShapeType.GEAR_9),
    HALF_FRAME(STShapeType.HALF_FRAME),
    HEART(STShapeType.HEART),
    HEPTAGON(STShapeType.HEPTAGON),
    HEXAGON(STShapeType.HEXAGON),
    HOME_PLATE(STShapeType.HOME_PLATE),
    HORIZONTAL_SCROLL(STShapeType.HORIZONTAL_SCROLL),
    IRREGULAR_SEAL_1(STShapeType.IRREGULAR_SEAL_1),
    IRREGULAR_SEAL_2(STShapeType.IRREGULAR_SEAL_2),
    LEFT_ARROW(STShapeType.LEFT_ARROW),
    LEFT_ARROW_CALLOUT(STShapeType.LEFT_ARROW_CALLOUT),
    LEFT_BRACE(STShapeType.LEFT_BRACE),
    LEFT_BRACKET(STShapeType.LEFT_BRACKET),
    LEFT_CIRCULAR_ARROW(STShapeType.LEFT_CIRCULAR_ARROW),
    LEFT_RIGHT_ARROW(STShapeType.LEFT_RIGHT_ARROW),
    LEFT_RIGHT_ARROW_CALLOUT(STShapeType.LEFT_RIGHT_ARROW_CALLOUT),
    LEFT_RIGHT_CIRCULAR_ARROW(STShapeType.LEFT_RIGHT_CIRCULAR_ARROW),
    LEFT_RIGHT_RIBBON(STShapeType.LEFT_RIGHT_RIBBON),
    LEFT_RIGHT_UP_ARROW(STShapeType.LEFT_RIGHT_UP_ARROW),
    LEFT_UP_ARROW(STShapeType.LEFT_UP_ARROW),
    LIGHTNING_BOLT(STShapeType.LIGHTNING_BOLT),
    LINE(STShapeType.LINE),
    LINE_INVERTED(STShapeType.LINE_INV),
    MATH_DIVIDE(STShapeType.MATH_DIVIDE),
    MATH_EQUAL(STShapeType.MATH_EQUAL),
    MATH_MINUS(STShapeType.MATH_MINUS),
    MATH_MULTIPLY(STShapeType.MATH_MULTIPLY),
    MATH_NOT_EQUAL(STShapeType.MATH_NOT_EQUAL),
    MATH_PLUS(STShapeType.MATH_PLUS),
    MOON(STShapeType.MOON),
    NO_SMOKING(STShapeType.NO_SMOKING),
    NON_ISOSCELES_TRAPEZOID(STShapeType.NON_ISOSCELES_TRAPEZOID),
    NOTCHED_RIGHT_ARROW(STShapeType.NOTCHED_RIGHT_ARROW),
    OCTAGON(STShapeType.OCTAGON),
    PARALLELOGRAM(STShapeType.PARALLELOGRAM),
    PENTAGON(STShapeType.PENTAGON),
    PIE(STShapeType.PIE),
    PIE_WEDGE(STShapeType.PIE_WEDGE),
    PLAQUE(STShapeType.PLAQUE),
    PLAQUE_TABS(STShapeType.PLAQUE_TABS),
    PLUS(STShapeType.PLUS),
    QUAD_ARROW(STShapeType.QUAD_ARROW),
    QUAD_ARROW_CALLOUT(STShapeType.QUAD_ARROW_CALLOUT),
    RECTANGLE(STShapeType.RECT),
    RIBBON(STShapeType.RIBBON),
    RIBBON_2(STShapeType.RIBBON_2),
    RIGHT_ARROW(STShapeType.RIGHT_ARROW),
    RIGHT_ARROW_CALLOUT(STShapeType.RIGHT_ARROW_CALLOUT),
    RIGHT_BRACE(STShapeType.RIGHT_BRACE),
    RIGHT_BRACKET(STShapeType.RIGHT_BRACKET),
    ROUND_RECTANGLE_1_CORNER(STShapeType.ROUND_1_RECT),
    ROUND_RECTANGLE_2_DIAGONAL_CORNERS(STShapeType.ROUND_2_DIAG_RECT),
    ROUND_RECTANGLE_2_SAME_SIDE_CORNERS(STShapeType.ROUND_2_SAME_RECT),
    ROUND_RECTANGLE(STShapeType.ROUND_RECT),
    RIGHT_TRIANGLE(STShapeType.RT_TRIANGLE),
    SMILEY_FACE(STShapeType.SMILEY_FACE),
    SNIP_RECTANGLE_1_CORNER(STShapeType.SNIP_1_RECT),
    SNIP_RECTANGLE_2_DIAGONAL_CORNERS(STShapeType.SNIP_2_DIAG_RECT),
    SNIP_RECTANGLE_2_SAME_SIDE_CORNERS(STShapeType.SNIP_2_SAME_RECT),
    SNIP_ROUND_RECTANGLE(STShapeType.SNIP_ROUND_RECT),
    SQUARE_TABS(STShapeType.SQUARE_TABS),
    STAR_10(STShapeType.STAR_10),
    STAR_12(STShapeType.STAR_12),
    STAR_16(STShapeType.STAR_16),
    STAR_24(STShapeType.STAR_24),
    STAR_32(STShapeType.STAR_32),
    STAR_4(STShapeType.STAR_4),
    STAR_5(STShapeType.STAR_5),
    STAR_6(STShapeType.STAR_6),
    STAR_7(STShapeType.STAR_7),
    STAR_8(STShapeType.STAR_8),
    STRAIGHT_CONNECTOR(STShapeType.STRAIGHT_CONNECTOR_1),
    STRIPED_RIGHT_ARROW(STShapeType.STRIPED_RIGHT_ARROW),
    SUN(STShapeType.SUN),
    SWOOSH_ARROW(STShapeType.SWOOSH_ARROW),
    TEARDROP(STShapeType.TEARDROP),
    TRAPEZOID(STShapeType.TRAPEZOID),
    TRIANGLE(STShapeType.TRIANGLE),
    UP_ARROW(STShapeType.UP_ARROW),
    UP_ARROW_CALLOUT(STShapeType.UP_ARROW_CALLOUT),
    UP_DOWN_ARROW(STShapeType.UP_DOWN_ARROW),
    UP_DOWN_ARROW_CALLOUT(STShapeType.UP_DOWN_ARROW_CALLOUT),
    UTURN_ARROW(STShapeType.UTURN_ARROW),
    VERTICAL_SCROLL(STShapeType.VERTICAL_SCROLL),
    WAVE(STShapeType.WAVE),
    WEDGE_ELLIPSE_CALLOUT(STShapeType.WEDGE_ELLIPSE_CALLOUT),
    WEDGE_RECTANGLE_CALLOUT(STShapeType.WEDGE_RECT_CALLOUT),
    WEDGE_ROUND_RECTANGLE_CALLOUT(STShapeType.WEDGE_ROUND_RECT_CALLOUT);

    final STShapeType.Enum underlying;

    PresetGeometry(STShapeType.Enum shape) {
        this.underlying = shape;
    }

    private static final HashMap<STShapeType.Enum, PresetGeometry> reverse = new HashMap<>();
    static {
        for (PresetGeometry value : values()) {
            reverse.put(value.underlying, value);
        }
    }

    static PresetGeometry valueOf(STShapeType.Enum shape) {
        return reverse.get(shape);
    }
}
