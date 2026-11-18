// The item-kind matrix, set `mix`, shard 2 of 5 (sharded to keep each process small).

import { matrixShard } from './run.ts';

matrixShard('mix', 1, 5);
