<?php
/*Bibliotecas */
//utiliza bibliotecas do composer
require_once 'vendor/autoload.php';
require_once 'assets/class/LibPhpPresentationManipulation.php';
//bibliotecas envolvidas



/*Constantes de apoio*/
//files
const IMAGE_STORAGE = __DIR__ . './assets/images';//diretorio das imagens
const PRESENTATION_STORAGE = __DIR__ . './assets/presentation_files'; //diretorio dos arquivos de apresentacao
//style
const TITLE_PRIMARY_COLOR = "FFFFFF";

/*Instancia da manipulacao*/
$lib_pptx = new LibPhpPresentationManipulation();

/*Criacao */
//inicia nova apresentacao
$presentation = $lib_pptx::new_presentation();
//cria slide
$slide_1 = $lib_pptx::new_slide($presentation);

// Slide > Background > Image
$lib_pptx->set_background_image_in_slide(
    $slide_1, //slide
    "ibs-bg.png" //nome do arquivo
);

//titulo_capa
$lib_pptx->create_text(
    $slide_1, //slide
    300, //altura
    600, //largura
    20, //posicao no eixo X
    400, //posicao no eixo Y
    "Relatório de Imprensa", //texto
    true, //ativo bold
    40, //font size
    TITLE_PRIMARY_COLOR //cor da fonte
);
//subtiulo

//salva arquivo
$lib_pptx->create_pptx_file(
    $presentation, //apresentacao
    'PowerPoint2007', //tipo de arquivo
    "teste-ibs.pptx"
);

