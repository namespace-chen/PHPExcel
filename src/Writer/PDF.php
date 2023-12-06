<?php

namespace PHPExcel\Writer;



/**
 *  \PHPExcel\Writer\PDF
 *
 *  Copyright (c) 2006 - 2015 PHPExcel
 *
 *  This library is free software; you can redistribute it and/or
 *  modify it under the terms of the GNU Lesser General Public
 *  License as published by the Free Software Foundation; either
 *  version 2.1 of the License, or (at your option) any later version.
 *
 *  This library is distributed in the hope that it will be useful,
 *  but WITHOUT ANY WARRANTY; without even the implied warranty of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 *  Lesser General Public License for more details.
 *
 *  You should have received a copy of the GNU Lesser General Public
 *  License along with this library; if not, write to the Free Software
 *  Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 *  @category    PHPExcel
 *  @package     \PHPExcel\Writer\PDF
 *  @copyright   Copyright (c) 2006 - 2015 PHPExcel (http://www.codeplex.com/PHPExcel)
 *  @license     http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 *  @version     ##VERSION##, ##DATE##
 */
class PDF implements \PHPExcel\Writer\IWriter
{

   /**
    * The wrapper for the requested PDF rendering engine
    *
    * @var \PHPExcel\Writer\PDF\Core
    */
   private $renderer = null;

   /**
    *  Instantiate a new renderer of the configured type within this container class
    *
    *  @param \PHPExcel\PHPExcel   $phpExcel         PHPExcel object
    *  @throws \PHPExcel\Writer\Exception    when PDF library is not configured
    */
   public function __construct(\PHPExcel\PHPExcel $phpExcel)
   {
      $pdfLibraryName = \PHPExcel\Settings::getPdfRendererName();
      if (is_null($pdfLibraryName)) {
         throw new \PHPExcel\Writer\Exception("PDF Rendering library has not been defined.");
      }

      $pdfLibraryPath = \PHPExcel\Settings::getPdfRendererPath();
      if (is_null($pdfLibraryName)) {
         throw new \PHPExcel\Writer\Exception("PDF Rendering library path has not been defined.");
      }
      $includePath = str_replace('\\', '/', get_include_path());
      $rendererPath = str_replace('\\', '/', $pdfLibraryPath);
      if (strpos($rendererPath, $includePath) === false) {
         set_include_path(get_include_path() . PATH_SEPARATOR . $pdfLibraryPath);
      }

      $rendererName = "\\PHPExcel\\Writer\\PDF\\{$pdfLibraryName}";
      $this->renderer = new $rendererName($phpExcel);
   }


   /**
    *  Magic method to handle direct calls to the configured PDF renderer wrapper class.
    *
    *  @param   string   $name        Renderer library method name
    *  @param   mixed[]  $arguments   Array of arguments to pass to the renderer method
    *  @return  mixed    Returned data from the PDF renderer wrapper method
    */
   public function __call($name, $arguments)
   {
      if ($this->renderer === null) {
         throw new \PHPExcel\Writer\Exception("PDF Rendering library has not been defined.");
      }

      return call_user_func_array(array($this->renderer, $name), $arguments);
   }

   /**
    * {@inheritdoc}
    */
   public function save($pFilename = null)
   {
      $this->renderer->save($pFilename);
   }
}
