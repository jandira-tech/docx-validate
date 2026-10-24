/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * The default registry of named checks. Registration order is dependency
 * order: a check appears after every check it requires.
 */

import { CheckRegistry } from "../../../lib/check-registry";
import { zipBytesAfterEocdCheck, zipEocdCommentOverrunCheck } from "./package";
import { wordRowWithoutCellCheck } from "./tables";
import { xmlSyntaxCheck } from "./xml";

export function createDefaultCheckRegistry(): CheckRegistry {
    return new CheckRegistry([zipBytesAfterEocdCheck, zipEocdCommentOverrunCheck, xmlSyntaxCheck, wordRowWithoutCellCheck]);
}
