package net.cgrand;

import java.io.File;
import java.io.InputStream;

import clojure.lang.IPersistentMap;
import clojure.lang.Var;

public class SpreadMap {
	static final private Var spreadmap =
			clojure.lang.RT.var("net.cgrand.spreadmap", "spreadmap");
	
	static IPersistentMap create(File file) {
		return (IPersistentMap) spreadmap.invoke(file);
	};
	
	static IPersistentMap create(String filename) {
		return (IPersistentMap) spreadmap.invoke(filename);
	};
	
	/**
	 * Consumes (and close) the provided inputstream to create a spreadmap.
	 * @param input
	 * @return a spreadmap
	 */
	static IPersistentMap create(InputStream input) {
		return (IPersistentMap) spreadmap.invoke(input);
	};
}
