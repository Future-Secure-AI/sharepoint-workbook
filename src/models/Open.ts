import type { DriveItemRef } from "microsoft-graph/DriveItem";

export type OpenRef = {
	id: OpenId;
	itemRef?: DriveItemRef;
};

export type OpenId = string & {
	readonly __brand: unique symbol;
};
