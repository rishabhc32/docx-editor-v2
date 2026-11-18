// The item-kind matrix, set `nested`, shard 3 of 3 (sharded to keep each process small).

import { matrixShard } from './run.ts';

matrixShard('nested', 2, 3);
