import { defineConfig } from "vitest/config";

export default defineConfig({
	test: {
		environment: "node",
		watch: false,
		include: ["src/**/*.test.ts"],
		testTimeout: 16000,
	},
});
