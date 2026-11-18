// The item-kind matrix, set `pairs`, shard 2 of 2 (sharded to keep each process small).

import { matrixShard } from './run.ts';

matrixShard('pairs', 1, 2);
