<?php

/*Libs relacionadas*/
use PhpOffice\PhpPresentation\Style\Alignment; //classe de estilo de alinhamentos
use PhpOffice\PhpPresentation\Style\Color; //classe de estilo de cores
use PhpOffice\PhpPresentation\PhpPresentation; //classe do PhpPresentation
use PhpOffice\PhpPresentation\Slide\Background\Image; //utilizacao de imagens
use PhpOffice\PhpPresentation\IOFactory; //classe para manipular os arquivos



class LibPhpPresentationManipulation
{
    //desc: define a criacao de box como text
    //params: (obj) getActiveSlide, (string) tipo do box
    //return: (obj) createRichTextShape
    static public function type_box($slide, $type){
        switch ($type){
            case 'RichTextShape':
                return $slide->createRichTextShape();
                break;
        }
    }

    //desc: cracao de box para o texto
    //params: (string) type box, (obj) getActiveSlide, (number) altura, (number) largura, (number) posicao eixo X, (number) posicao eixo y
    //return: (obj) createRichTextShape
    static function create_box($created_box, $height, $width, $offsetX, $offsetY){
        //espaco ocupado pela forma
        $created_box->setHeight($height); //altura
        $created_box->setWidth($width); //largura
        $created_box->setOffsetX($offsetX); //posicao em relacao ao eixo X
        $created_box->setOffsetY($offsetY); //posicao em relacao ao eixo Y
        //retorna a box apos formacao
        return $created_box;
    }

    //desc: criacao de texto
    //params: (obj) getActiveSlide, (number) altura, (number) largura, (number) posicao eixo X, (number) posicao eixo y, () alinhamento, (string) texto, (bool) se bold, (number) fonte0size, (string) color
    //return: (obj) RichTextShape
    static public function create_text($created_box, $alignment, $text, $isBold, $fontSize, $color){
        //cria box de texto
        $title = $created_box;
        //texto
        //se alinhamento do paragrafo (centralizado)
        if($alignment == "HORIZONTAL_CENTER")$title->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        if($alignment == "HORIZONTAL_LEFT")$title->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
        //texto do paragrafo
        $text = $title->createTextRun($text);
        //fonte do texto como negrito
        $text->getFont()->setBold($isBold)
            ->setSize($fontSize) //tamanho da fonte
            ->setColor(new Color($color)); //cor da fonte
    }

    //desc: cria apresentacao
    //params: nenehum
    //return: (obj) PhpPresentation
    static function new_presentation(){
        return new PhpPresentation();
    }

    //desc: cria slide
    //params: (obj) PhpPresentation, (bool) se primeiro slide
    //return: (obj) getActiveSlide
    static function new_slide($presentation, $first_slide = false){
        //se for o primeiro slide
        if($first_slide) return $presentation->getActiveSlide();
        //demais
        return $presentation->createSlide();
    }

    //desc: cria imagem
    //params: (obj) getActiveSlide, (string) caminho de imagem
    //return: nenhum
    static function set_background_image_in_slide($slide, $file_name){
        $bg_image = new Image();
        $bg_image->setPath(IMAGE_STORAGE."/$file_name");
        $slide->setBackground($bg_image);
    }

    //desc: cria arquivo
    //params: (obj) PhpPresentation, (string) tipo de arquivo, (string) nome do arquivo
    //return
    static function create_pptx_file($presentation, $file_type, $file_name){
        $oWriterPPTX = IOFactory::createWriter($presentation, $file_type); //definindo o tipo de arquivo como PowerPoint2007
        $oWriterPPTX->save(PRESENTATION_STORAGE. "/$file_name");
    }
}