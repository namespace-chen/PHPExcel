<?php

namespace PHPExcel;


/**
 * \PHPExcel\Chart
 *
 * Copyright (c) 2006 - 2015 PHPExcel
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category    PHPExcel
 * @package        \PHPExcel\Chart
 * @copyright    Copyright (c) 2006 - 2015 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license        http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version        ##VERSION##, ##DATE##
 */
class Chart
{
   /**
    * Chart Name
    *
    * @var string
    */
   private $name = '';

   /**
    * Worksheet
    *
    * @var \PHPExcel\Worksheet
    */
   private $worksheet;

   /**
    * Chart Title
    *
    * @var \PHPExcel\Chart\Title
    */
   private $title;

   /**
    * Chart Legend
    *
    * @var \PHPExcel\Chart\Legend
    */
   private $legend;

   /**
    * X-Axis Label
    *
    * @var \PHPExcel\Chart\Title
    */
   private $xAxisLabel;

   /**
    * Y-Axis Label
    *
    * @var \PHPExcel\Chart\Title
    */
   private $yAxisLabel;

   /**
    * Chart Plot Area
    *
    * @var \PHPExcel\Chart\PlotArea
    */
   private $plotArea;

   /**
    * Plot Visible Only
    *
    * @var boolean
    */
   private $plotVisibleOnly = true;

   /**
    * Display Blanks as
    *
    * @var string
    */
   private $displayBlanksAs = '0';

   /**
    * Chart Asix Y as
    *
    * @var \PHPExcel\Chart\Axis
    */
   private $yAxis;

   /**
    * Chart Asix X as
    *
    * @var \PHPExcel\Chart\Axis
    */
   private $xAxis;

   /**
    * Chart Major Gridlines as
    *
    * @var \PHPExcel\Chart\GridLines
    */
   private $majorGridlines;

   /**
    * Chart Minor Gridlines as
    *
    * @var \PHPExcel\Chart\GridLines
    */
   private $minorGridlines;

   /**
    * Top-Left Cell Position
    *
    * @var string
    */
   private $topLeftCellRef = 'A1';


   /**
    * Top-Left X-Offset
    *
    * @var integer
    */
   private $topLeftXOffset = 0;


   /**
    * Top-Left Y-Offset
    *
    * @var integer
    */
   private $topLeftYOffset = 0;


   /**
    * Bottom-Right Cell Position
    *
    * @var string
    */
   private $bottomRightCellRef = 'A1';


   /**
    * Bottom-Right X-Offset
    *
    * @var integer
    */
   private $bottomRightXOffset = 10;


   /**
    * Bottom-Right Y-Offset
    *
    * @var integer
    */
   private $bottomRightYOffset = 10;


   /**
    * Create a new \PHPExcel\Chart
    */
   public function __construct($name, \PHPExcel\Chart\Title $title = null, \PHPExcel\Chart\Legend $legend = null, \PHPExcel\Chart\PlotArea $plotArea = null, $plotVisibleOnly = true, $displayBlanksAs = '0', \PHPExcel\Chart\Title $xAxisLabel = null, \PHPExcel\Chart\Title $yAxisLabel = null, \PHPExcel\Chart\Axis $xAxis = null, \PHPExcel\Chart\Axis $yAxis = null, \PHPExcel\Chart\GridLines $majorGridlines = null, \PHPExcel\Chart\GridLines $minorGridlines = null)
   {
      $this->name = $name;
      $this->title = $title;
      $this->legend = $legend;
      $this->xAxisLabel = $xAxisLabel;
      $this->yAxisLabel = $yAxisLabel;
      $this->plotArea = $plotArea;
      $this->plotVisibleOnly = $plotVisibleOnly;
      $this->displayBlanksAs = $displayBlanksAs;
      $this->xAxis = $xAxis;
      $this->yAxis = $yAxis;
      $this->majorGridlines = $majorGridlines;
      $this->minorGridlines = $minorGridlines;
   }

   /**
    * Get Name
    *
    * @return string
    */
   public function getName()
   {
      return $this->name;
   }

   /**
    * Get Worksheet
    *
    * @return \PHPExcel\Worksheet
    */
   public function getWorksheet()
   {
      return $this->worksheet;
   }

   /**
    * Set Worksheet
    *
    * @param    \PHPExcel\Worksheet    $pValue
    * @throws    \PHPExcel\Chart\Exception
    * @return \PHPExcel\Chart
    */
   public function setWorksheet(\PHPExcel\Worksheet $pValue = null)
   {
      $this->worksheet = $pValue;

      return $this;
   }

   /**
    * Get Title
    *
    * @return \PHPExcel\Chart\Title
    */
   public function getTitle()
   {
      return $this->title;
   }

   /**
    * Set Title
    *
    * @param    \PHPExcel\Chart\Title $title
    * @return    \PHPExcel\Chart
    */
   public function setTitle(\PHPExcel\Chart\Title $title)
   {
      $this->title = $title;

      return $this;
   }

   /**
    * Get Legend
    *
    * @return \PHPExcel\Chart\Legend
    */
   public function getLegend()
   {
      return $this->legend;
   }

   /**
    * Set Legend
    *
    * @param    \PHPExcel\Chart\Legend $legend
    * @return    \PHPExcel\Chart
    */
   public function setLegend(\PHPExcel\Chart\Legend $legend)
   {
      $this->legend = $legend;

      return $this;
   }

   /**
    * Get X-Axis Label
    *
    * @return \PHPExcel\Chart\Title
    */
   public function getXAxisLabel()
   {
      return $this->xAxisLabel;
   }

   /**
    * Set X-Axis Label
    *
    * @param    \PHPExcel\Chart\Title $label
    * @return    \PHPExcel\Chart
    */
   public function setXAxisLabel(\PHPExcel\Chart\Title $label)
   {
      $this->xAxisLabel = $label;

      return $this;
   }

   /**
    * Get Y-Axis Label
    *
    * @return \PHPExcel\Chart\Title
    */
   public function getYAxisLabel()
   {
      return $this->yAxisLabel;
   }

   /**
    * Set Y-Axis Label
    *
    * @param    \PHPExcel\Chart\Title $label
    * @return    \PHPExcel\Chart
    */
   public function setYAxisLabel(\PHPExcel\Chart\Title $label)
   {
      $this->yAxisLabel = $label;

      return $this;
   }

   /**
    * Get Plot Area
    *
    * @return \PHPExcel\Chart\PlotArea
    */
   public function getPlotArea()
   {
      return $this->plotArea;
   }

   /**
    * Get Plot Visible Only
    *
    * @return boolean
    */
   public function getPlotVisibleOnly()
   {
      return $this->plotVisibleOnly;
   }

   /**
    * Set Plot Visible Only
    *
    * @param boolean $plotVisibleOnly
    * @return \PHPExcel\Chart
    */
   public function setPlotVisibleOnly($plotVisibleOnly = true)
   {
      $this->plotVisibleOnly = $plotVisibleOnly;

      return $this;
   }

   /**
    * Get Display Blanks as
    *
    * @return string
    */
   public function getDisplayBlanksAs()
   {
      return $this->displayBlanksAs;
   }

   /**
    * Set Display Blanks as
    *
    * @param string $displayBlanksAs
    * @return \PHPExcel\Chart
    */
   public function setDisplayBlanksAs($displayBlanksAs = '0')
   {
      $this->displayBlanksAs = $displayBlanksAs;
   }


   /**
    * Get yAxis
    *
    * @return \PHPExcel\Chart\Axis
    */
   public function getChartAxisY()
   {
      if ($this->yAxis !== null) {
         return $this->yAxis;
      }

      return new \PHPExcel\Chart\Axis();
   }

   /**
    * Get xAxis
    *
    * @return \PHPExcel\Chart\Axis
    */
   public function getChartAxisX()
   {
      if ($this->xAxis !== null) {
         return $this->xAxis;
      }

      return new \PHPExcel\Chart\Axis();
   }

   /**
    * Get Major Gridlines
    *
    * @return \PHPExcel\Chart\GridLines
    */
   public function getMajorGridlines()
   {
      if ($this->majorGridlines !== null) {
         return $this->majorGridlines;
      }

      return new \PHPExcel\Chart\GridLines();
   }

   /**
    * Get Minor Gridlines
    *
    * @return \PHPExcel\Chart\GridLines
    */
   public function getMinorGridlines()
   {
      if ($this->minorGridlines !== null) {
         return $this->minorGridlines;
      }

      return new \PHPExcel\Chart\GridLines();
   }


   /**
    * Set the Top Left position for the chart
    *
    * @param    string    $cell
    * @param    integer    $xOffset
    * @param    integer    $yOffset
    * @return \PHPExcel\Chart
    */
   public function setTopLeftPosition($cell, $xOffset = null, $yOffset = null)
   {
      $this->topLeftCellRef = $cell;
      if (!is_null($xOffset)) {
         $this->setTopLeftXOffset($xOffset);
      }
      if (!is_null($yOffset)) {
         $this->setTopLeftYOffset($yOffset);
      }

      return $this;
   }

   /**
    * Get the top left position of the chart
    *
    * @return array    an associative array containing the cell address, X-Offset and Y-Offset from the top left of that cell
    */
   public function getTopLeftPosition()
   {
      return array(
         'cell'    => $this->topLeftCellRef,
         'xOffset' => $this->topLeftXOffset,
         'yOffset' => $this->topLeftYOffset
      );
   }

   /**
    * Get the cell address where the top left of the chart is fixed
    *
    * @return string
    */
   public function getTopLeftCell()
   {
      return $this->topLeftCellRef;
   }

   /**
    * Set the Top Left cell position for the chart
    *
    * @param    string    $cell
    * @return \PHPExcel\Chart
    */
   public function setTopLeftCell($cell)
   {
      $this->topLeftCellRef = $cell;

      return $this;
   }

   /**
    * Set the offset position within the Top Left cell for the chart
    *
    * @param    integer    $xOffset
    * @param    integer    $yOffset
    * @return \PHPExcel\Chart
    */
   public function setTopLeftOffset($xOffset = null, $yOffset = null)
   {
      if (!is_null($xOffset)) {
         $this->setTopLeftXOffset($xOffset);
      }
      if (!is_null($yOffset)) {
         $this->setTopLeftYOffset($yOffset);
      }

      return $this;
   }

   /**
    * Get the offset position within the Top Left cell for the chart
    *
    * @return integer[]
    */
   public function getTopLeftOffset()
   {
      return array(
         'X' => $this->topLeftXOffset,
         'Y' => $this->topLeftYOffset
      );
   }

   public function setTopLeftXOffset($xOffset)
   {
      $this->topLeftXOffset = $xOffset;

      return $this;
   }

   public function getTopLeftXOffset()
   {
      return $this->topLeftXOffset;
   }

   public function setTopLeftYOffset($yOffset)
   {
      $this->topLeftYOffset = $yOffset;

      return $this;
   }

   public function getTopLeftYOffset()
   {
      return $this->topLeftYOffset;
   }

   /**
    * Set the Bottom Right position of the chart
    *
    * @param    string    $cell
    * @param    integer    $xOffset
    * @param    integer    $yOffset
    * @return \PHPExcel\Chart
    */
   public function setBottomRightPosition($cell, $xOffset = null, $yOffset = null)
   {
      $this->bottomRightCellRef = $cell;
      if (!is_null($xOffset)) {
         $this->setBottomRightXOffset($xOffset);
      }
      if (!is_null($yOffset)) {
         $this->setBottomRightYOffset($yOffset);
      }

      return $this;
   }

   /**
    * Get the bottom right position of the chart
    *
    * @return array    an associative array containing the cell address, X-Offset and Y-Offset from the top left of that cell
    */
   public function getBottomRightPosition()
   {
      return array(
         'cell'    => $this->bottomRightCellRef,
         'xOffset' => $this->bottomRightXOffset,
         'yOffset' => $this->bottomRightYOffset
      );
   }

   public function setBottomRightCell($cell)
   {
      $this->bottomRightCellRef = $cell;

      return $this;
   }

   /**
    * Get the cell address where the bottom right of the chart is fixed
    *
    * @return string
    */
   public function getBottomRightCell()
   {
      return $this->bottomRightCellRef;
   }

   /**
    * Set the offset position within the Bottom Right cell for the chart
    *
    * @param    integer    $xOffset
    * @param    integer    $yOffset
    * @return \PHPExcel\Chart
    */
   public function setBottomRightOffset($xOffset = null, $yOffset = null)
   {
      if (!is_null($xOffset)) {
         $this->setBottomRightXOffset($xOffset);
      }
      if (!is_null($yOffset)) {
         $this->setBottomRightYOffset($yOffset);
      }

      return $this;
   }

   /**
    * Get the offset position within the Bottom Right cell for the chart
    *
    * @return integer[]
    */
   public function getBottomRightOffset()
   {
      return array(
         'X' => $this->bottomRightXOffset,
         'Y' => $this->bottomRightYOffset
      );
   }

   public function setBottomRightXOffset($xOffset)
   {
      $this->bottomRightXOffset = $xOffset;

      return $this;
   }

   public function getBottomRightXOffset()
   {
      return $this->bottomRightXOffset;
   }

   public function setBottomRightYOffset($yOffset)
   {
      $this->bottomRightYOffset = $yOffset;

      return $this;
   }

   public function getBottomRightYOffset()
   {
      return $this->bottomRightYOffset;
   }


   public function refresh()
   {
      if ($this->worksheet !== null) {
         $this->plotArea->refresh($this->worksheet);
      }
   }

   public function render($outputDestination = null)
   {
      $libraryName = \PHPExcel\Settings::getChartRendererName();
      if (is_null($libraryName)) {
         return false;
      }
      //    Ensure that data series values are up-to-date before we render
      $this->refresh();

      $libraryPath = \PHPExcel\Settings::getChartRendererPath();
      $includePath = str_replace('\\', '/', get_include_path());
      $rendererPath = str_replace('\\', '/', $libraryPath);
      if (strpos($rendererPath, $includePath) === false) {
         set_include_path(get_include_path() . PATH_SEPARATOR . $libraryPath);
      }

      $rendererName = '\PHPExcel\Chart\Renderer\' . $libraryName;
      $renderer = new $rendererName($this);

      if ($outputDestination == 'php://output') {
         $outputDestination = null;
      }
      return $renderer->render($outputDestination);
   }
}
