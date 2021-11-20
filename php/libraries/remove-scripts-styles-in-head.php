<?php

/* Remove head tag contents */
/* Remove action */
function remove_scripts_styles_in_head(){
  if( !is_admin() ){
		remove_action( 'wp_head', 'wp_print_scripts' );
		remove_action( 'wp_head', 'wp_print_head_scripts', 9 );
		remove_action( 'wp_head', 'wp_enqueue_scripts', 1 );
	  remove_action( 'wp_head', 'www-widgetapi' );
	  remove_action( 'wp_head', 'wp_generator' );
	  remove_action( 'wp_head', 'rsd_link' );
	  remove_action( 'wp_head', 'wlwmanifest_link' );
	  remove_action( 'wp_head', 'index_rel_link' );
	  remove_action( 'wp_head', 'wp_shortlink_wp_head', 10, 0 );
	  remove_action( 'wp_head', 'parent_post_rel_link', 10, 0 );
	  remove_action( 'wp_head', 'start_post_rel_link', 10, 0 );
	  remove_action( 'wp_head', 'adjacent_posts_rel_link_wp_head', 10, 0 );
	  remove_action( 'wp_head', 'feed_links', 2 );
	  remove_action( 'wp_head', 'feed_links_extra', 3 );
	  remove_action( 'wp_head', 'print_emoji_detection_script', 7 );
	  remove_action( 'admin_print_scripts', 'print_emoji_detection_script' );
	  remove_action( 'wp_print_styles', 'print_emoji_styles' );
	  remove_action( 'admin_print_styles', 'print_emoji_styles' );

		add_filter( 'emoji_svg_url', '__return_false' );
  }
}

add_action('wp_enqueue_scripts', 'remove_scripts_styles_in_head');

 ?>
