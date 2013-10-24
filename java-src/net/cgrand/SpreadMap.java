package net.cgrand;

import java.io.File;
import java.io.InputStream;

import clojure.lang.Associative;
import clojure.lang.Var;

public class SpreadMap {
	static {
		clojure.lang.RT.var("clojure.core", "require")
		.invoke(clojure.lang.Symbol.create("net.cgrand.spreadmap"));
	}
	
	static final private Var spreadmap =
			clojure.lang.RT.var("net.cgrand.spreadmap", "spreadmap");
	
	static public Associative create(File file) {
		return (Associative) spreadmap.invoke(file);
	};
	
	static public Associative create(String filename) {
		return (Associative) spreadmap.invoke(filename);
	};
	
	/**
	 * Consumes (and close) the provided inputstream to create a spreadmap.
	 * @param input
	 * @return a spreadmap
	 */
	static public Associative create(InputStream input) {
		return (Associative) spreadmap.invoke(input);
	};
}
