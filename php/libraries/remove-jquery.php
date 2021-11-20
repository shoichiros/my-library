<?php

/* remove jQuery */
function remove_jquery( $scripts ){
    if( !is_admin() ){
        $scripts->remove( 'jquery' );
    }
}

add_filter( 'wp_default_scripts', 'remove_jquery' );


 ?>
