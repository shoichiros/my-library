<?php

/* Make shortcode */
function automatically_link_and_title( $atts ) {
  $content_atts = shortcode_atts( array(
      'id' => '',
  ), $atts );

	$href_link = get_the_permalink( $content_atts['id'] );
	$title = get_the_title( $content_atts['id'] );
	$link = "<a href=" . $href_link . ">" . $title . "</a>";

	return $link;
}

add_shortcode( 'link_auto', 'automatically_link_and_title' );

?>
