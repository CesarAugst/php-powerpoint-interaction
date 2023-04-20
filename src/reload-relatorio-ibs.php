<?php
/*Bibliotecas */
//utiliza bibliotecas do composer
require_once 'vendor/autoload.php';
require_once 'assets/class/LibPhpPresentationManipulation.php';

/*Constantes de apoio*/
//files
const IMAGE_STORAGE = __DIR__ . './assets/images';//diretorio das imagens
const PRESENTATION_STORAGE = __DIR__ . './assets/presentation_files'; //diretorio dos arquivos de apresentacao
//style
const TITLE_PRIMARY_COLOR = "FFFFFF";
const TITLE_SECONDARY_COLOR = "FF000000";
//alignment
const HORIZONTAL_CENTER = "HORIZONTAL_CENTER";
const HORIZONTAL_LEFT = "HORIZONTAL_LEFT";
//bullet
const TYPE_BULLET = "TYPE_BULLET";
const TYPE_NONE = "TYPE_NONE";
//type box
const RICHTEXTSHAPE = "RICHTEXTSHAPE";
const FILEIMAGE = "FILEIMAGE";
const BASE64IMAGE = "BASE64IMAGE";

/*Instancia da manipulacao*/
$lib_pptx = new LibPhpPresentationManipulation();

/*Criacao */
//inicia nova apresentacao
$presentation = $lib_pptx::load_presentation(
    "teste-ibs.pptx", //nome do arquivo
    'pptx' //versao do arquivo (pptx ou odp)
);

//cria slide
$slide_2 = $lib_pptx::new_slide($presentation);
// define o background
$slide_base = $presentation->getAllSlides()[0];
$lib_pptx::set_existing_background(
    $slide_2, //slide alvo
    $slide_base //slide base
);
//cria box para o titulo
$created_box = $lib_pptx::create_box(
    $lib_pptx::type_box($slide_2, RICHTEXTSHAPE), //tipo de box
    300, //altura
    600, //largura
    300, //posicao no eixo X
    300 //posicao no eixo Y
);
//titulo
$lib_pptx::create_text(
    $created_box, //box do texto
    $lib_pptx::type_alignment(HORIZONTAL_LEFT), //alinhamento do texto
    "Obrigado", //texto
    true, //ativo bold
    50, //font size
    TITLE_PRIMARY_COLOR //cor da fonte
);

//salva arquivo
$lib_pptx::create_pptx_file(
    $presentation, //apresentacao
    'PowerPoint2007', //tipo de arquivo
    "teste-reload-ibs.pptx"
);

