package com.template.engine.utils;

import java.util.regex.Pattern;

public class DocxPattern {

	public static final Pattern FIELD_PATTERN_1 = Pattern.compile("\\$\\{field:[a-zA-Z]+\\}");
	public static final Pattern FIELD_PATTERN_2 = Pattern.compile("\\$\\{field:[a-zA-Z]+\\.[a-zA-Z]+\\}");
	
	public static final Pattern IMAGE_PATTERN = Pattern.compile("\\$\\{image:[a-zA-Z]+\\}");
	
	public static final Pattern HEADER_PATTERN = Pattern.compile("\\$\\{header:[a-zA-Z]+\\}");
	public static final Pattern HEADER_PATTERN_2 = Pattern.compile("\\$\\{header:[a-zA-Z]+\\.[a-zA-Z]+\\}");
	
	public static final Pattern FOOTER_PATTERN = Pattern.compile("\\$\\{footer:[a-zA-Z]+\\}");
	public static final Pattern FOOTER_PATTERN_2 = Pattern.compile("\\$\\{footer:[a-zA-Z]+\\.[a-zA-Z]+\\}");
	
	public static final Pattern OBJECT_FIELD_PATTERN_1 = Pattern.compile("[a-zA-Z]+\\.[a-zA-Z]+:[a-zA-Z]+");
	public static final Pattern OBJECT_FIELD_PATTERN_2 = Pattern.compile("[a-zA-Z]+:[a-zA-Z]+");
	public static final Pattern OBJECT_FIELD_PATTERN_3 = Pattern.compile("[a-zA-Z]+\\.[a-zA-Z]+");
	
	public static final Pattern COLLECTION_START_PATTERN_1 = Pattern.compile("\\$\\{collection:[a-zA-Z]+:[a-zA-Z]+\\}");
	public static final Pattern COLLECTION_START_PATTERN_2 = Pattern.compile("\\$\\{collection:[a-zA-Z]+\\.[a-zA-Z]+:[a-zA-Z]+\\}");
	public static final Pattern COLLECTION_OBJECT_PATTERN = Pattern.compile("\\$\\{[a-zA-Z]+\\.[a-zA-Z]+\\}");
	public static final Pattern COLLECTION_END_PATTERN_1 = Pattern.compile("\\$\\{/collection:[a-zA-Z]+:[a-zA-Z]+\\}");
	public static final Pattern COLLECTION_END_PATTERN_2 = Pattern.compile("\\$\\{/collection:[a-zA-Z]+\\.[a-zA-Z]+:[a-zA-Z]+\\}");
	
	
}
