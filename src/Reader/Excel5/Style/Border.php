<?php

namespace PHPExcel\Reader\Excel5\Style;



class Border
{
   protected static $map = array(
      0x00 => \PHPExcel\Style\Border::BORDER_NONE,
      0x01 => \PHPExcel\Style\Border::BORDER_THIN,
      0x02 => \PHPExcel\Style\Border::BORDER_MEDIUM,
      0x03 => \PHPExcel\Style\Border::BORDER_DASHED,
      0x04 => \PHPExcel\Style\Border::BORDER_DOTTED,
      0x05 => \PHPExcel\Style\Border::BORDER_THICK,
      0x06 => \PHPExcel\Style\Border::BORDER_DOUBLE,
      0x07 => \PHPExcel\Style\Border::BORDER_HAIR,
      0x08 => \PHPExcel\Style\Border::BORDER_MEDIUMDASHED,
      0x09 => \PHPExcel\Style\Border::BORDER_DASHDOT,
      0x0A => \PHPExcel\Style\Border::BORDER_MEDIUMDASHDOT,
      0x0B => \PHPExcel\Style\Border::BORDER_DASHDOTDOT,
      0x0C => \PHPExcel\Style\Border::BORDER_MEDIUMDASHDOTDOT,
      0x0D => \PHPExcel\Style\Border::BORDER_SLANTDASHDOT,
   );

   /**
    * Map border style
    * OpenOffice documentation: 2.5.11
    *
    * @param int $index
    * @return string
    */
   public static function lookup($index)
   {
      if (isset(self::$map[$index])) {
         return self::$map[$index];
      }
      return \PHPExcel\Style\Border::BORDER_NONE;
   }
}
