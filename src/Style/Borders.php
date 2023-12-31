<?php

namespace PHPExcel\Style;



/**
 * \PHPExcel\Style\Borders
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
 * @category   PHPExcel
 * @package    \PHPExcel\Style
 * @copyright  Copyright (c) 2006 - 2015 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
class Borders extends \PHPExcel\Style\Supervisor implements \PHPExcel\IComparable
{
   /* Diagonal directions */
   const DIAGONAL_NONE = 0;
   const DIAGONAL_UP   = 1;
   const DIAGONAL_DOWN = 2;
   const DIAGONAL_BOTH = 3;

   /**
    * Left
    *
    * @var \PHPExcel\Style\Border
    */
   protected $left;

   /**
    * Right
    *
    * @var \PHPExcel\Style\Border
    */
   protected $right;

   /**
    * Top
    *
    * @var \PHPExcel\Style\Border
    */
   protected $top;

   /**
    * Bottom
    *
    * @var \PHPExcel\Style\Border
    */
   protected $bottom;

   /**
    * Diagonal
    *
    * @var \PHPExcel\Style\Border
    */
   protected $diagonal;

   /**
    * DiagonalDirection
    *
    * @var int
    */
   protected $diagonalDirection;

   /**
    * All borders psedo-border. Only applies to supervisor.
    *
    * @var \PHPExcel\Style\Border
    */
   protected $allBorders;

   /**
    * Outline psedo-border. Only applies to supervisor.
    *
    * @var \PHPExcel\Style\Border
    */
   protected $outline;

   /**
    * Inside psedo-border. Only applies to supervisor.
    *
    * @var \PHPExcel\Style\Border
    */
   protected $inside;

   /**
    * Vertical pseudo-border. Only applies to supervisor.
    *
    * @var \PHPExcel\Style\Border
    */
   protected $vertical;

   /**
    * Horizontal pseudo-border. Only applies to supervisor.
    *
    * @var \PHPExcel\Style\Border
    */
   protected $horizontal;

   /**
    * Create a new \PHPExcel\Style\Borders
    *
    * @param    boolean    $isSupervisor    Flag indicating if this is a supervisor or not
    *                                    Leave this value at default unless you understand exactly what
    *                                        its ramifications are
    * @param    boolean    $isConditional    Flag indicating if this is a conditional style or not
    *                                    Leave this value at default unless you understand exactly what
    *                                        its ramifications are
    */
   public function __construct($isSupervisor = false, $isConditional = false)
   {
      // Supervisor?
      parent::__construct($isSupervisor);

      // Initialise values
      $this->left = new \PHPExcel\Style\Border($isSupervisor, $isConditional);
      $this->right = new \PHPExcel\Style\Border($isSupervisor, $isConditional);
      $this->top = new \PHPExcel\Style\Border($isSupervisor, $isConditional);
      $this->bottom = new \PHPExcel\Style\Border($isSupervisor, $isConditional);
      $this->diagonal = new \PHPExcel\Style\Border($isSupervisor, $isConditional);
      $this->diagonalDirection = \PHPExcel\Style\Borders::DIAGONAL_NONE;

      // Specially for supervisor
      if ($isSupervisor) {
         // Initialize pseudo-borders
         $this->allBorders = new \PHPExcel\Style\Border(true);
         $this->outline = new \PHPExcel\Style\Border(true);
         $this->inside = new \PHPExcel\Style\Border(true);
         $this->vertical = new \PHPExcel\Style\Border(true);
         $this->horizontal = new \PHPExcel\Style\Border(true);

         // bind parent if we are a supervisor
         $this->left->bindParent($this, 'left');
         $this->right->bindParent($this, 'right');
         $this->top->bindParent($this, 'top');
         $this->bottom->bindParent($this, 'bottom');
         $this->diagonal->bindParent($this, 'diagonal');
         $this->allBorders->bindParent($this, 'allBorders');
         $this->outline->bindParent($this, 'outline');
         $this->inside->bindParent($this, 'inside');
         $this->vertical->bindParent($this, 'vertical');
         $this->horizontal->bindParent($this, 'horizontal');
      }
   }

   /**
    * Get the shared style component for the currently active cell in currently active sheet.
    * Only used for style supervisor
    *
    * @return \PHPExcel\Style\Borders
    */
   public function getSharedComponent()
   {
      return $this->parent->getSharedComponent()->getBorders();
   }

   /**
    * Build style array from subcomponents
    *
    * @param array $array
    * @return array
    */
   public function getStyleArray($array)
   {
      return array('borders' => $array);
   }

   /**
    * Apply styles from array
    *
    * <code>
    * $objPHPExcel->getActiveSheet()->getStyle('B2')->getBorders()->applyFromArray(
    *         array(
    *             'bottom'     => array(
    *                 'style' => \PHPExcel\Style\Border::BORDER_DASHDOT,
    *                 'color' => array(
    *                     'rgb' => '808080'
    *                 )
    *             ),
    *             'top'     => array(
    *                 'style' => \PHPExcel\Style\Border::BORDER_DASHDOT,
    *                 'color' => array(
    *                     'rgb' => '808080'
    *                 )
    *             )
    *         )
    * );
    * </code>
    * <code>
    * $objPHPExcel->getActiveSheet()->getStyle('B2')->getBorders()->applyFromArray(
    *         array(
    *             'allborders' => array(
    *                 'style' => \PHPExcel\Style\Border::BORDER_DASHDOT,
    *                 'color' => array(
    *                     'rgb' => '808080'
    *                 )
    *             )
    *         )
    * );
    * </code>
    *
    * @param    array    $pStyles    Array containing style information
    * @throws    \PHPExcel\Exception
    * @return \PHPExcel\Style\Borders
    */
   public function applyFromArray($pStyles = null)
   {
      if (is_array($pStyles)) {
         if ($this->isSupervisor) {
            $this->getActiveSheet()->getStyle($this->getSelectedCells())->applyFromArray($this->getStyleArray($pStyles));
         } else {
            if (array_key_exists('left', $pStyles)) {
               $this->getLeft()->applyFromArray($pStyles['left']);
            }
            if (array_key_exists('right', $pStyles)) {
               $this->getRight()->applyFromArray($pStyles['right']);
            }
            if (array_key_exists('top', $pStyles)) {
               $this->getTop()->applyFromArray($pStyles['top']);
            }
            if (array_key_exists('bottom', $pStyles)) {
               $this->getBottom()->applyFromArray($pStyles['bottom']);
            }
            if (array_key_exists('diagonal', $pStyles)) {
               $this->getDiagonal()->applyFromArray($pStyles['diagonal']);
            }
            if (array_key_exists('diagonaldirection', $pStyles)) {
               $this->setDiagonalDirection($pStyles['diagonaldirection']);
            }
            if (array_key_exists('allborders', $pStyles)) {
               $this->getLeft()->applyFromArray($pStyles['allborders']);
               $this->getRight()->applyFromArray($pStyles['allborders']);
               $this->getTop()->applyFromArray($pStyles['allborders']);
               $this->getBottom()->applyFromArray($pStyles['allborders']);
            }
         }
      } else {
         throw new \PHPExcel\Exception("Invalid style array passed.");
      }
      return $this;
   }

   /**
    * Get Left
    *
    * @return \PHPExcel\Style\Border
    */
   public function getLeft()
   {
      return $this->left;
   }

   /**
    * Get Right
    *
    * @return \PHPExcel\Style\Border
    */
   public function getRight()
   {
      return $this->right;
   }

   /**
    * Get Top
    *
    * @return \PHPExcel\Style\Border
    */
   public function getTop()
   {
      return $this->top;
   }

   /**
    * Get Bottom
    *
    * @return \PHPExcel\Style\Border
    */
   public function getBottom()
   {
      return $this->bottom;
   }

   /**
    * Get Diagonal
    *
    * @return \PHPExcel\Style\Border
    */
   public function getDiagonal()
   {
      return $this->diagonal;
   }

   /**
    * Get AllBorders (pseudo-border). Only applies to supervisor.
    *
    * @return \PHPExcel\Style\Border
    * @throws \PHPExcel\Exception
    */
   public function getAllBorders()
   {
      if (!$this->isSupervisor) {
         throw new \PHPExcel\Exception('Can only get pseudo-border for supervisor.');
      }
      return $this->allBorders;
   }

   /**
    * Get Outline (pseudo-border). Only applies to supervisor.
    *
    * @return boolean
    * @throws \PHPExcel\Exception
    */
   public function getOutline()
   {
      if (!$this->isSupervisor) {
         throw new \PHPExcel\Exception('Can only get pseudo-border for supervisor.');
      }
      return $this->outline;
   }

   /**
    * Get Inside (pseudo-border). Only applies to supervisor.
    *
    * @return boolean
    * @throws \PHPExcel\Exception
    */
   public function getInside()
   {
      if (!$this->isSupervisor) {
         throw new \PHPExcel\Exception('Can only get pseudo-border for supervisor.');
      }
      return $this->inside;
   }

   /**
    * Get Vertical (pseudo-border). Only applies to supervisor.
    *
    * @return \PHPExcel\Style\Border
    * @throws \PHPExcel\Exception
    */
   public function getVertical()
   {
      if (!$this->isSupervisor) {
         throw new \PHPExcel\Exception('Can only get pseudo-border for supervisor.');
      }
      return $this->vertical;
   }

   /**
    * Get Horizontal (pseudo-border). Only applies to supervisor.
    *
    * @return \PHPExcel\Style\Border
    * @throws \PHPExcel\Exception
    */
   public function getHorizontal()
   {
      if (!$this->isSupervisor) {
         throw new \PHPExcel\Exception('Can only get pseudo-border for supervisor.');
      }
      return $this->horizontal;
   }

   /**
    * Get DiagonalDirection
    *
    * @return int
    */
   public function getDiagonalDirection()
   {
      if ($this->isSupervisor) {
         return $this->getSharedComponent()->getDiagonalDirection();
      }
      return $this->diagonalDirection;
   }

   /**
    * Set DiagonalDirection
    *
    * @param int $pValue
    * @return \PHPExcel\Style\Borders
    */
   public function setDiagonalDirection($pValue = \PHPExcel\Style\Borders::DIAGONAL_NONE)
   {
      if ($pValue == '') {
         $pValue = \PHPExcel\Style\Borders::DIAGONAL_NONE;
      }
      if ($this->isSupervisor) {
         $styleArray = $this->getStyleArray(array('diagonaldirection' => $pValue));
         $this->getActiveSheet()->getStyle($this->getSelectedCells())->applyFromArray($styleArray);
      } else {
         $this->diagonalDirection = $pValue;
      }
      return $this;
   }

   /**
    * Get hash code
    *
    * @return string    Hash code
    */
   public function getHashCode()
   {
      if ($this->isSupervisor) {
         return $this->getSharedComponent()->getHashcode();
      }
      return md5(
         $this->getLeft()->getHashCode() .
            $this->getRight()->getHashCode() .
            $this->getTop()->getHashCode() .
            $this->getBottom()->getHashCode() .
            $this->getDiagonal()->getHashCode() .
            $this->getDiagonalDirection() .
            __CLASS__
      );
   }
}
