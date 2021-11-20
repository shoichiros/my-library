<?php

/* Quick tag remove and add */
/* leave quicktags */

function leave_quicktags( $qtInit ) {
	$qtInit['buttons'] = 'link,ul,ol,li,fullscreen';
	return $qtInit;
}

add_filter('quicktags_settings', 'leave_quicktags');

/* Add again admin editor quicktags */
/* QTags.addButton( id, show_name, start_tag, end_tag ); */

function add_quicktags() {
	if( wp_script_is('quicktags') ) {
?>
	<script type="text/javascript">
	  QTags.addButton( 'mystrong', 'strong', '<strong>', '</strong>' );
		QTags.addButton( 'h1', 'h1', '<h1>', '</h1>' );
		QTags.addButton( 'h2', 'h2', '<h2>', '</h2>' );
		QTags.addButton( 'h3', 'h3', '<h3>', '</h3>' );
	</script>
<?php
	}
}

add_action( 'admin_print_footer_scripts', 'add_quicktags' );


 ?>
