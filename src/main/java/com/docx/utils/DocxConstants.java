package com.docx.utils;

public class DocxConstants {

	public static final String DEFAULT_TAG_START = "${";
	public static final String DEFAULT_TAG_END = "}";
	
	public static final String TAG_PREFIX_FIELD = "field:";
	public static final String TAG_PREFIX_IMAGE = "image:";
	public static final String TAG_PREFIX_HEADER = "header:";
	public static final String TAG_PREFIX_FOOTER = "footer:";
	public static final String TAG_PREFIX_COLLECTION_START = "collection:";
	public static final String TAG_PREFIX_COLLECTION_END = "/collection:";
	
	public static final String EMPTY_STRING = "";
	
	private DocxConstants() {
	    throw new IllegalStateException("Constants class");
	  }
}
