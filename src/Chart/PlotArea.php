<?php

namespace PHPExcel\Chart;



/**
 * \PHPExcel\Chart\PlotArea
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
class PlotArea
{
   /**
    * PlotArea Layout
    *
    * @var \PHPExcel\Chart\Layout
    */
   private $layout = null;

   /**
    * Plot Series
    *
    * @var array of \PHPExcel\Chart\DataSeries
    */
   private $plotSeries = array();

   /**
    * Create a new \PHPExcel\Chart\PlotArea
    */
   public function __construct(\PHPExcel\Chart\Layout $layout = null, $plotSeries = array())
   {
      $this->layout = $layout;
      $this->plotSeries = $plotSeries;
   }

   /**
    * Get Layout
    *
    * @return \PHPExcel\Chart\Layout
    */
   public function getLayout()
   {
      return $this->layout;
   }

   /**
    * Get Number of Plot Groups
    *
    * @return array of \PHPExcel\Chart\DataSeries
    */
   public function getPlotGroupCount()
   {
      return count($this->plotSeries);
   }

   /**
    * Get Number of Plot Series
    *
    * @return integer
    */
   public function getPlotSeriesCount()
   {
      $seriesCount = 0;
      foreach ($this->plotSeries as $plot) {
         $seriesCount += $plot->getPlotSeriesCount();
      }
      return $seriesCount;
   }

   /**
    * Get Plot Series
    *
    * @return array of \PHPExcel\Chart\DataSeries
    */
   public function getPlotGroup()
   {
      return $this->plotSeries;
   }

   /**
    * Get Plot Series by Index
    *
    * @return \PHPExcel\Chart\DataSeries
    */
   public function getPlotGroupByIndex($index)
   {
      return $this->plotSeries[$index];
   }

   /**
    * Set Plot Series
    *
    * @param [\PHPExcel\Chart\DataSeries]
    * @return \PHPExcel\Chart\PlotArea
    */
   public function setPlotSeries($plotSeries = array())
   {
      $this->plotSeries = $plotSeries;

      return $this;
   }

   public function refresh(\PHPExcel\Worksheet $worksheet)
   {
      foreach ($this->plotSeries as $plotSeries) {
         $plotSeries->refresh($worksheet);
      }
   }
}
