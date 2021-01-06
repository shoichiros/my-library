<?php

/* Custom menu register */
/* wp_nav_menu($arrayname = array('theme_location' => 'left name', 'menu' => 'right name' );) */
function register_custom_menu() {
	/* Custom menu header, footer and aside */
	register_nav_menus(array(
		'footer_menu' => __( 'Footer Menu' ),
	));
}

add_action('init', 'register_custom_menu');


 ?>
