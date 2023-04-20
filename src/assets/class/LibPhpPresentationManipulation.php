<?php

/*Libs relacionadas*/
use PhpOffice\PhpPresentation\Style\Alignment; //classe de estilo de alinhamentos
use PhpOffice\PhpPresentation\Style\Color; //classe de estilo de cores
use PhpOffice\PhpPresentation\PhpPresentation; //classe do PhpPresentation
use PhpOffice\PhpPresentation\Slide\Background\Image; //utilizacao de imagens
use PhpOffice\PhpPresentation\IOFactory; //classe para manipular os arquivos
use PhpOffice\PhpPresentation\Style\Bullet; //Bullet
use PhpOffice\PhpPresentation\Shape\Drawing; //desenhos



class LibPhpPresentationManipulation
{
    //desc: cracao de box para o texto
    //params: (string) type box, (obj) getActiveSlide, (number) altura, (number) largura, (number) posicao eixo X, (number) posicao eixo y
    //return: (obj) createRichTextShape
    static function create_box($created_type_box, $height, $width, $offsetX, $offsetY){
        //espaco ocupado pela forma
        $created_type_box->setHeight($height); //altura
        $created_type_box->setWidth($width); //largura
        $created_type_box->setOffsetX($offsetX); //posicao em relacao ao eixo X
        $created_type_box->setOffsetY($offsetY); //posicao em relacao ao eixo Y
        //retorna a box apos formacao
        return $created_type_box;
    }

    //desc: criacao de texto
    //params: (obj) box criada, (obj) alinhamento, (string) texto, (bool) se bold, (number) fonte0size, (string) color
    //return: (obj) RichTextShape
    static public function create_text($created_box, $alignment, $text, $isBold, $fontSize, $color){
        //cria box de texto
        $shape = $created_box;
        //alinhamento
        $shape->getActiveParagraph()->getAlignment()->setHorizontal($alignment);
        //texto do paragrafo
        $textRun = $shape->createTextRun($text);
        //fonte do texto como negrito
        $textRun->getFont()->setBold($isBold)
            ->setSize($fontSize) //tamanho da fonte
            ->setColor(new Color($color)); //cor da fonte
        //retorna o texto
        return $textRun;
    }

    //desc: criacao de texto com paragrafos
    //params: (obj) getActiveSlide, (obj) alinhamento, (array<string>) texto, (bool) se bold, (number) fonte0size, (string) color, (obj) bullet
    //return: (obj) RichTextShape
    static public function create_paragraph_text($created_box, $alignment_type, $array_text, $isBold, $fontSize, $color, $bullet_type){
        //cria box de texto
        $shape = $created_box;
        //alinhamento
        $shape->getActiveParagraph()->getAlignment()->setHorizontal($alignment_type)
            ->setMarginLeft(25)
            ->setIndent(-25);
        //fonte do texto como negrito
        $shape->getActiveParagraph()->getFont()->setBold($isBold) //negrito
            ->setSize($fontSize) //tamanho da fonte
            ->setColor(new Color($color)); //cor da fonte
        $shape->getActiveParagraph()->getBulletStyle()->setBulletType($bullet_type);
        //texto
        foreach ($array_text as $index => $text){
            if($index == 0){
                $shape->createTextRun($text);
            }else{
                $shape->createParagraph()->createTextRun($text);
            }
            $shape->createBreak();
        }
        //retorna o paragrafo
        return $shape;
    }

    //desc: criacao de imagem
    //params:
    //return:
    static public function create_image($created_box, $file_name, $image_type){
        $created_box
            ->setName('NOME') //nome da imagem
            ->setDescription('DESCRICAO'); //descricao da imagem
        if($image_type == FILEIMAGE){
            return $created_box->setPath(IMAGE_STORAGE."/$file_name"); //caminho para a imagem
        }else{
            $img_64 = base64_encode(file_get_contents($file_name));// conversao da imagem
            return $created_box->setData("data:image/jpeg;base64,$img_64"); //incorporando imagem
        }

    }

    //desc: define o bullet do texto
    //params: (string) tipo do alinhamento
    //return: (obj) Alignment
    static public function type_bullet($type){
        switch ($type){
            case 'TYPE_BULLET':
                return Bullet::TYPE_BULLET;
                break;
            case 'NONE':
            default:
                return Bullet::TYPE_NONE;
        }
    }

    //desc: define a criacao de box como text
    //params: (obj) getActiveSlide, (string) tipo do box
    //return: (obj) createRichTextShape
    static public function type_box($slide, $type){
        switch ($type){
            case 'FILEIMAGE':
                $shape = new Drawing\File();
                return $slide->addShape($shape);
                break;
            case 'BASE64IMAGE':
                $shape = new Drawing\Base64();
                return $slide->addShape($shape);
            case 'RICHTEXTSHAPE':
            default:
                return $slide->createRichTextShape();
                break;
        }
    }

    //desc: define o alinhamento do texto
    //params: (string) tipo do alinhamento
    //return: (obj) Alignment
    static public function type_alignment($type){
        switch ($type){
            case 'HORIZONTAL_CENTER':
                return Alignment::HORIZONTAL_CENTER;
                break;
            case 'HORIZONTAL_LEFT':
            default:
                return Alignment::HORIZONTAL_LEFT;
        }
    }

    //desc: cria apresentacao
    //params: nenehum
    //return: (obj) PhpPresentation
    static function new_presentation(){
        return new PhpPresentation();
    }

    //carrega apresentacao ja criada
    //params: (string) nome da apresentacao
    //return: (obj) PhpPresentation
    static function load_presentation($file_name, $version_file){
        //relaciona versao com extensao
        $extension_by_version = $version_file == 'pptx' ? 'PowerPoint2007' : 'ODPresentation';
        //abre a leitura do arquivo no tipo indicado
        $reader = IOFactory::createReader($extension_by_version);
        return $reader->load(PRESENTATION_STORAGE."/$file_name");
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

    //desc: define como background do slide um background de outro ja existente
    //params: (obj) getActiveSlide, (obj) slide
    //return: nenhum
    static function set_existing_background($slide_target, $slide_base){
        //pega o background do slide base
        $background = $slide_base->getBackground();
        //define no slide alvo o atual background
        $slide_target->setBackground($background);
    }

    //desc: cria arquivo
    //params: (obj) PhpPresentation, (string) tipo de arquivo, (string) nome do arquivo
    //return
    static function create_pptx_file($presentation, $file_type, $file_name){
        $oWriterPPTX = IOFactory::createWriter($presentation, $file_type); //definindo o tipo de arquivo como PowerPoint2007
        $oWriterPPTX->save(PRESENTATION_STORAGE. "/$file_name");
    }
}