<?php

namespace PHPExcel\Reader\Excel5\Style;



class FillPattern
{
   protected static $map = array(
      0x00 => \PHPExcel\Style\Fill::FILL_NONE,
      0x01 => \PHPExcel\Style\Fill::FILL_SOLID,
      0x02 => \PHPExcel\Style\Fill::FILL_PATTERN_MEDIUMGRAY,
      0x03 => \PHPExcel\Style\Fill::FILL_PATTERN_DARKGRAY,
      0x04 => \PHPExcel\Style\Fill::FILL_PATTERN_LIGHTGRAY,
      0x05 => \PHPExcel\Style\Fill::FILL_PATTERN_DARKHORIZONTAL,
      0x06 => \PHPExcel\Style\Fill::FILL_PATTERN_DARKVERTICAL,
      0x07 => \PHPExcel\Style\Fill::FILL_PATTERN_DARKDOWN,
      0x08 => \PHPExcel\Style\Fill::FILL_PATTERN_DARKUP,
      0x09 => \PHPExcel\Style\Fill::FILL_PATTERN_DARKGRID,
      0x0A => \PHPExcel\Style\Fill::FILL_PATTERN_DARKTRELLIS,
      0x0B => \PHPExcel\Style\Fill::FILL_PATTERN_LIGHTHORIZONTAL,
      0x0C => \PHPExcel\Style\Fill::FILL_PATTERN_LIGHTVERTICAL,
      0x0D => \PHPExcel\Style\Fill::FILL_PATTERN_LIGHTDOWN,
      0x0E => \PHPExcel\Style\Fill::FILL_PATTERN_LIGHTUP,
      0x0F => \PHPExcel\Style\Fill::FILL_PATTERN_LIGHTGRID,
      0x10 => \PHPExcel\Style\Fill::FILL_PATTERN_LIGHTTRELLIS,
      0x11 => \PHPExcel\Style\Fill::FILL_PATTERN_GRAY125,
      0x12 => \PHPExcel\Style\Fill::FILL_PATTERN_GRAY0625,
   );

   /**
    * Get fill pattern from index
    * OpenOffice documentation: 2.5.12
    *
    * @param int $index
    * @return string
    */
   public static function lookup($index)
   {
      if (isset(self::$map[$index])) {
         return self::$map[$index];
      }
      return \PHPExcel\Style\Fill::FILL_NONE;
   }
}
