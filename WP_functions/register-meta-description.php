<?php

/* Meta description on home, singlepost and category */
/* classic editor show excerpt text box */
function register_meta_description() {
	global $post;

	if ( is_singular() ) {
		$post_description = strip_tags( $post->post_excerpt );
		$post_description = mb_substr( $post_description, 0, 140, 'utf-8' );
		echo '<meta name="description" content="' . $post_description . '">';
	}

	if( is_home() ) {
		echo '<meta name="description" content="' . get_bloginfo( 'description' ) . '">';
	}
  /* If making category.php, bottom code add */
  /*
	if( is_category() ) {
		$cat_description = strip_tags( category_description() );
		echo '<meta name="description" content="' . $cat_description . '">';
	}
  */
}

add_action( 'wp_head', 'register_meta_description' );

?>
