Made this util to export table directly into Excel.

## Features
- Streaming XLSX export
- Minimal memory usage
- Fast large dataset writes
- Bounded queue for backpressure control
- Manual XML writing (no DOM usage)
- No external dependencies; XLSX is generated via direct ZIP + XML streaming

## Data types
- STRING
- NUMBER
- BOOLEAN
- DATE

## Implementation notes
- Reused buffers for serialization
- Preallocated column metadata
- Streaming ZIP-based writer
- Producer-consumer pipeline

## Example
See `np.com.subasadhikari.simulations.WideSheetSimulation` for a streaming usage example with batching and throughput measurement.
