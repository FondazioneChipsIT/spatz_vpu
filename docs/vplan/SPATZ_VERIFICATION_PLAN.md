# Spatz Verification Plan

This document is the verification plan of the standalone Spatz vector unit (top-level module `spatz`, `hw/src/spatz.sv`).
It follows the Foundation *Verification Methodology*: Phase 1 (specification analysis) is written out below, Phase 2 (the plan itself) is the spreadsheet [SPATZ_VERIFICATION_PLAN.xlsx](SPATZ_VERIFICATION_PLAN.xlsx), and the last section collects the inputs for Phase 3 (environment build).

| Artifact | Content |
|----------|---------|
| [SPATZ_SPEC.md](../SPATZ_SPEC.md) | Specification of the DUT, reconstructed from the RTL (source of every *Note* quote) |
| This document | Phase 1 lists, configuration-space analysis, open questions, Phase 3 inputs |
| [SPATZ_VERIFICATION_PLAN.xlsx](SPATZ_VERIFICATION_PLAN.xlsx) | Phase 2 plan: sheets *Verification Plan* and *Configuration* |
| [gen_spatz_vplan.py](gen_spatz_vplan.py) | Generator of the spreadsheet; aborts if a *Note* is not a verbatim quote of its source |

The spreadsheet is the place where the status of the verification is read. Regenerate it with `python3 docs/vplan/gen_spatz_vplan.py` from `working_dir/spatz_vpu` (needs `openpyxl`); once the verification engineers start updating *Status* and *Test Name* by hand, edit the spreadsheet directly and stop regenerating it.

**Table of Contents**
- [1. Phase 1 — Specification analysis](#1-phase-1--specification-analysis)
  - [1.1 Where the specification came from](#11-where-the-specification-came-from)
  - [1.2 Interfaces](#12-interfaces)
  - [1.3 Functional requirements](#13-functional-requirements)
  - [1.4 Configuration parameters](#14-configuration-parameters)
  - [1.5 Assumptions about the environment](#15-assumptions-about-the-environment)
  - [1.6 Open questions](#16-open-questions)
  - [1.7 Known limitations](#17-known-limitations)
- [2. Phase 2 — Verification plan](#2-phase-2--verification-plan)
  - [2.1 How the spreadsheet is filled](#21-how-the-spreadsheet-is-filled)
  - [2.2 Old testbench](#22-old-testbench)
  - [2.3 Configuration space](#23-configuration-space)
  - [2.4 A worked item](#24-a-worked-item)
- [3. Inputs for Phase 3 — Environment build](#3-inputs-for-phase-3--environment-build)


## 1. Phase 1 — Specification analysis

### 1.1 Where the specification came from

Spatz ships without a written specification of its top level. [SPATZ_SPEC.md](../SPATZ_SPEC.md) was reconstructed from the source (§3.1 of the methodology), in this order of confidence:

1. the RTL of `working_dir/spatz_vpu/hw/src` — port lists, parameters, FSMs, decode tables;
2. the adopted protocols — TCDM channels (`tcdm_interface`) for the VLSU ports and `reqrsp_interface` for the FP LSU;
3. the intended usage — the reference integration `spatz_cc.sv` / `spatz_cluster.sv`, the Snitch core that drives the offload interface, the self-checking tests in `sw/riscvTests` and `sw/spatzBenchmarks`;
4. the upstream README and commit history.

The RVV 1.0 specification is the reference for instruction semantics, restricted to the subset listed in *Supported Instructions* of the spec and in the instruction spreadsheet linked there.
Behaviours that could only be read from the RTL, and that no protocol, usage or explicit design decision justifies, are **not** requirements: they are listed as [open questions](#16-open-questions). Requirements that rest on such a reading are tagged `[assumed]` and collected in [Known limitations](#17-known-limitations).

### 1.2 Interfaces

Single clock domain (`clk_i`, rising edge) and single asynchronous active-low reset (`rst_ni`). Widths refer to the default configuration (`spatz_cluster.default.dram.hjson`).

| Interface | Ports | Protocol | Direction of initiative |
|-----------|-------|----------|-------------------------|
| Offload issue | `issue_valid_i`, `issue_ready_o`, `issue_req_i`, `issue_rsp_o` | valid/ready; `issue_rsp_o` combinational, valid in the handshake cycle (except `isfloat`, level) | core → DUT |
| Offload write-back | `rsp_valid_o`, `rsp_ready_i`, `rsp_o` | valid/ready, registered outputs | DUT → core |
| VLSU memory ×`NrMemPorts` (4) | `spatz_mem_req_o/valid_o/ready_i`, `spatz_mem_rsp_i/valid_i` | TCDM channels: valid/ready request, response without ready, in order per port, one response per request | DUT → memory |
| FP LSU memory | `fp_lsu_mem_req_o`, `fp_lsu_mem_rsp_i` | `reqrsp`: independent valid/ready `q` and `p` channels, in order | DUT → memory |
| Memory completion | `spatz_mem_finished_o[1:0]`, `spatz_mem_str_finished_o[1:0]` | single-cycle pulses | DUT → core |
| FPU side channel | `fpu_rnd_mode_i`, `fpu_fmt_mode_i`, `fpu_status_o` | untimed inputs sampled at acceptance; status updated every cycle, OR-accumulated by the core | both |
| Miscellaneous | `testmode_i`, `hart_id_i` | static | integrator → DUT |

Details, data types and handshake rules: [SPATZ_SPEC.md — Top-Level Interface](../SPATZ_SPEC.md#top-level-interface). MemPool-only ports are out of scope.

### 1.3 Functional requirements

One entry per observable behaviour. The Feature ID is the unique identifier of the requirement; each feature is covered by one plan item of the same number. The full definition (goal, stimulus, verbatim spec quote, old testbench coverage, configurations) is in the *Verification Plan* sheet.

| Feature | Item | Requirement | Check | Status |
|---------|------|-------------|-------|--------|
| F000 | T000 | Offload issue handshake | assertion + scoreboard | Not Started |
| F001 | T001 | Issue immediate response | scoreboard | Not Started |
| F002 | T002 | isfloat busy indication | assertion + scoreboard | Not Started |
| F003 | T003 | Write-back response channel | scoreboard + assertion | In Progress |
| F004 | T004 | FP-destination results kept internal | assertion + scoreboard | In Progress |
| F005 | T005 | VLSU memory port protocol | assertion + scoreboard | In Progress |
| F006 | T006 | VLSU outstanding-request limits [assumed] | assertion + directed test | Not Started |
| F007 | T007 | FP LSU (reqrsp) interface | assertion + scoreboard | In Progress |
| F008 | T008 | Memory completion side channel | assertion + scoreboard | Not Started |
| F009 | T009 | Scalar FP vs vector memory ordering | scoreboard + directed test | Not Started |
| F010 | T010 | FPU side channel sampling | scoreboard + directed test | Not Started |
| F011 | T011 | Vector CSRs | scoreboard + directed test | In Progress |
| F012 | T012 | vsetvl* configuration | scoreboard | In Progress |
| F013 | T013 | vstart handling | scoreboard + directed test | Not Started |
| F014 | T014 | Masking (vm = 0) | scoreboard | In Progress |
| F015 | T015 | Tail policy [assumed] | scoreboard | Not Started |
| F016 | T016 | Integer vector arithmetic | scoreboard | In Progress |
| F017 | T017 | Widening integer arithmetic | scoreboard | In Progress |
| F018 | T018 | Integer multiply, multiply-add and divide | scoreboard | In Progress |
| F019 | T019 | Add/subtract with carry [open question Q03] | scoreboard | Not Started |
| F020 | T020 | Integer reductions | scoreboard | In Progress |
| F021 | T021 | Compares and mask-logical instructions | scoreboard | In Progress |
| F022 | T022 | Unit-stride loads and stores | scoreboard + assertion | In Progress |
| F023 | T023 | Strided loads and stores | scoreboard | In Progress |
| F024 | T024 | Indexed loads and stores | scoreboard | In Progress |
| F025 | T025 | Segmented loads and stores | scoreboard | In Progress |
| F026 | T026 | Permutation (VSLDU) | scoreboard | In Progress |
| F027 | T027 | FP vector arithmetic | scoreboard | In Progress |
| F028 | T028 | FP conversions | scoreboard | In Progress |
| F029 | T029 | FP widening and dot product [open question Q05] | scoreboard | In Progress |
| F030 | T030 | FP reductions [open question Q06] | scoreboard | In Progress |
| F031 | T031 | FP exception flags | scoreboard | Not Started |
| F032 | T032 | Scalar instructions executed by Spatz | scoreboard | In Progress |
| F033 | T033 | Illegal instructions [open questions Q01, Q02] | scoreboard + directed test | In Progress |
| F034 | T034 | Dependencies and chaining | scoreboard + constrained-random regression | In Progress |
| F035 | T035 | VRF bank arbitration and BUF_FPU | scoreboard | In Progress |
| F036 | T036 | Reset | assertion + directed test | Not Started |
| F037 | T037 | Liveness under backpressure | assertion + constrained-random regression | Not Started |
| F038 | T038 | Parameter rules | directed test (elaboration) + review | Not Started |
| F039 | T039 | Configuration space | constrained-random regression | In Progress |
| F040 | T040 | DOUBLE_BW VLSU | scoreboard | In Progress |
| F041 | T041 | Ventaglio extension | scoreboard | In Progress |

Groups: F000–F010 interfaces, F011–F015 programming model, F016–F032 instructions, F033 error handling, F034–F037 micro-architecture and infrastructure, F038–F041 parameters and optional configurations.

### 1.4 Configuration parameters

The parameters, with type, definition, legal range, default value and hjson knob, are in the *Configuration* sheet (derived from [SPATZ_SPEC.md — Parameters](../SPATZ_SPEC.md#parameters) and [Configuration](../SPATZ_SPEC.md#configuration)). Three sources drive them and nothing checks that they agree: the generated `spatz_pkg` (micro-architectural sizes), the top-level parameters set by the instantiating module, and the compile-time defines.

### 1.5 Assumptions about the environment

What Spatz is allowed to expect from the outside. In the real system they are guaranteed by the Snitch core and the TCDM; a standalone environment must enforce them (or deliberately violate them in dedicated negative tests only).

| ID | Assumption | Guaranteed in the cluster by |
|----|------------|------------------------------|
| A01 | `issue_req_i` is stable while `issue_valid_i` is high and `issue_ready_o` is low | Snitch offload logic |
| A02 | `issue_req_i.id[5] = 0`; `data_arga/argb` carry the current values of `rs1/rs2` (the core does not issue an instruction reading an integer register still pending in Spatz) | Snitch scoreboard |
| A03 | Scalar loads/stores of the core wait for Spatz memory instructions (stores wait for all of them, loads wait for Spatz stores), using `issue_rsp_o.loadstore` and `spatz_mem_*finished_o` | Snitch `acc_mem_cnt` counters |
| A04 | The core does not read `fflags`/`frm` while `issue_rsp_o.isfloat = 1` | Snitch `acc_stall` |
| A05 | `fpu_rnd_mode_i` / `fpu_fmt_mode_i` reflect `fcsr` and are stable in the handshake cycle of an FP instruction | Snitch `fcsr` |
| A06 | VLSU memory responses are returned in request order per port, one per request including stores, and are never back-pressured | TCDM interconnect |
| A07 | At most 7 stores per VLSU port are left unacknowledged `[assumed]` | TCDM latency of 1–2 cycles |
| A08 | The memory data width equals `ELEN` and the address width is 32 bits | cluster typedefs |
| A09 | FP LSU memory follows `reqrsp`: in-order responses, one per request | `reqrsp_mux` / `reqrsp_to_tcdm` |
| A10 | `FPUImplementation` is set explicitly (the declaration default disables every FPnew unit) | cluster wrapper |
| A11 | Every vector register is written before being read (the VRF is not reset) | software |

### 1.6 Open questions

Everything the specification does not answer. Each must be closed **in writing** with the designer or the architect: fill *Resolution* and change *Status*; the affected plan items are updated accordingly.

| ID | Question | Origin (RTL) | Plan items | Owner | Status | Resolution |
|----|----------|--------------|------------|-------|--------|------------|
| Q01 | `issue_rsp_o` reports `exception = 1` for vector instructions issued with `vill = 1`, but dispatch only checks the decoder `instr_illegal`: is the instruction meant to be dropped? It may currently be executed with an illegal `vtype`. | `spatz_controller.sv`, `ex_issue` / `acc_issue_resp` | T033 | designer | Open | |
| Q02 | `spatz_req_illegal` uses the current decoder output while the request is popped from the 1-entry buffer: an illegal instruction that cannot be popped in its decode cycle (e.g. response register full) could be executed later. Intended? | `spatz_controller.sv`, `ex_issue` | T033 | designer | Open | |
| Q03 | The VFU ties the IPU `carry_i` to `'0`: are `vadc`/`vsbc`/`vmadc.v*m`/`vmsbc.v*m` meant to take the carry/borrow from `v0`? | `spatz_vfu.sv`, `gen_ipus` | T019 | designer | Open | |
| Q04 | Tail is always undisturbed and `vtype.vta` is ignored: design decision (RVV allows it) or missing feature? | `spatz_vfu.sv`, `vreg_wbe_proc` | T015 | architect | Open | |
| Q05 | Widening FP operands are converted by exponent re-bias (`widen_fp*_to_fp*`) without handling zero, subnormals, infinities and NaNs: intended precision trade-off, and which reference behaviour must be checked? | `spatz_pkg.sv.tpl` | T029 | architect | Open | |
| Q06 | `vfredosum` is executed as a tree like `vfredusum`: accepted deviation from RVV ordered semantics? | `spatz_decoder.sv` | T030 | architect | Open | |
| Q07 | `csrrw` to `vl`/`vtype`/`vlenb` is silently ignored instead of being illegal: intended? | `spatz_decoder.sv`, CSR decode | T011 | designer | Open | |
| Q08 | `RegisterRsp` is unused and `NumOutstandingLoads` does not reach the VLSU (fixed at 8), while the cluster schema describes it as the VLSU buffer size: which is the intended behaviour? | `spatz.sv`, `spatz_controller.sv` | T039 | designer | Open | |
| Q09 | `fpu_status_o` is not qualified by the FPU output valid (sampled every cycle): can flags be reported late or twice, and is only the accumulated value architecturally meaningful? | `spatz_vfu.sv`, `gen_decoder` | T031 | designer | Open | |
| Q10 | FPnew `simd_mask_i` is tied to `'1`: are flags from masked-off and tail lanes expected in `fflags`? | `spatz_vfu.sv`, `gen_fpnew` | T031 | architect | Open | |
| Q11 | The declaration default `FPUImplementation = '0` disables every FPnew unit: should the module provide a working default? | `spatz.sv` | T038 | designer | Open | |
| Q12 | The VLSU store-acknowledge counter is 3 bits and stores are not throttled: is a memory with more than 7 outstanding stores per port legal (then this is a bug), or is A07 a documented integration constraint? | `spatz_vlsu.sv`, `store_count_q` | T006 | designer | Open | |
| Q13 | The strided address is `byte_index * (rs2 >> vsew)`: are strides that are not a multiple of `SEW/8` supported? | `spatz_vlsu.sv`, `gen_mem_req_addr` | T023 | designer | Open | |
| Q14 | `riscvTests-invalid` and `riscvTests-vill` are `WILL_FAIL`, but Snitch ignores `accept`/`exception`: through which mechanism do they fail, and what is the expected architectural outcome of an illegal offloaded instruction? | `sw/riscvTests/CMakeLists.txt`, `snitch.sv` | T033 | architect | Open | |
| Q15 | `double_bw` in the package and `` -DDOUBLE_BW `` must be set together, and the committed `generated/spatz_pkg.sv` has `double_bw: false`: is the standalone package expected to support both variants from one generation? | `spatz_pkg.sv.tpl`, `spatz.sv` | T040 | designer | Open | |

### 1.7 Known limitations

Requirements verified against an assumption are only as strong as the assumption:

- **F006 / T006 `[assumed]`** — the outstanding-request limits come from the RTL buffer sizes, not from a documented requirement (Q12, A07).
- **F015 / T015 `[assumed]`** — tail-undisturbed regardless of `vta` is the current RTL behaviour; it becomes a requirement only if Q04 confirms it.
- Items whose goal depends on Q01, Q02, Q03, Q05, Q06, Q13 check the behaviour stated in the spec; if an open question is resolved differently, the item and the spec are updated together.


## 2. Phase 2 — Verification plan

### 2.1 How the spreadsheet is filled

The *Verification Plan* sheet has one row per item with the columns of the methodology: *Item ID, Feature ID, Requirement, Verification Goal, Test Name, Stimulus, Status, Note, Old testbench situation, Configurations, Check*. Hovering a header shows how to fill the column.

- **Note** quotes the source verbatim (mostly SPATZ_SPEC.md, a few RTL comments); the generator refuses a quote that is not found in its source.
- **Test Name** lists the existing tests that already provide stimulus and check (`riscvTests-*`, `spatzBenchmarks`); `-` means no test exists yet.
- **Status** was set by the rule of §4.3: *In Progress* when the old regression provides stimulus and a check for part of the goal, *Not Started* otherwise. No item is *Passed*: the old regression runs at cluster level and never controls the DUT interfaces.
- **Configurations** refers to the configuration sets `CS-*` defined at the bottom of the *Configuration* sheet; empty means all legal configurations.
- A status summary (formulas) is below the last row.

### 2.2 Old testbench

The existing environment is the cluster-level software regression run by `util/local_ci.sh`:
- the `spatz_cluster` testbench (QuestaSim/VCS/Verilator) loads an ELF with `fesvr` and runs it on Snitch + Spatz;
- `sw/riscvTests` (115 enabled tests) and `sw/spatzBenchmarks` are self-checking C programs: results are stored to memory and compared in software with golden values (`VCMP`/`VMCMP` macros); pass/fail is the exit code;
- configurations: `default`, `smallvrf`, `32b`, `doublebw`, plus `ventaglio` for the benchmarks.

What it does **not** give, and therefore what every item's *Old testbench situation* column records as a gap:
- no control of the DUT interfaces: no random valid gaps, no backpressure on `rsp_ready_i`, fixed-latency memory with only natural bank conflicts;
- no protocol checks and no check of `issue_rsp_o`, `spatz_mem_*finished_o`, `fpu_status_o`;
- no randomisation and no functional coverage: in the enabled tests no one writes `vstart`, changes the rounding mode, reads `fflags` or uses fractional LMUL; masking (`v0.t`, 98/115 tests) and LMUL > 1 (101/115 tests) are exercised with fixed data.

### 2.3 Configuration space

Comparing the *Configuration* sheet with what the old regression runs makes the holes visible:

- **Swept:** `VLEN` (512, 256), ELEN/ISA (`rv32imafd`, `rv32imaf`), `double_bw`, `VENTAGLIO` (benchmarks only).
- **Fixed in every run — untested:** `N_FPU` = 4, `N_IPU` = 1 (a `4ipu` configuration exists but is not in `local_ci.sh`), `BUF_FPU` always defined, all FPU latencies (`timing.lat_*` identical in every configuration file), `fpu_pipe_config` = `BEFORE`, `NumOutstandingLoads` = 4, `RegisterRsp` = 1, `addr_width` = 32.
- **Correct by construction:** `NrMemPorts` (= `N_FU`), `NrVRFBanks`, `NrParallelInstructions`, FPnew `UnitTypes` (hard-coded), `data_width` (= `ELEN`).

Two new configuration sets are planned to close the most relevant holes: `CS-NOBUF` (`buf_fpu = 0`) and `CS-LAT` (different FPU latencies and pipeline placement); `CS-4IPU` must be added to the regression.

### 2.4 A worked item

| Field | Value |
|-------|-------|
| Item / Feature | T005 / F005 |
| Requirement | VLSU memory port protocol |
| Verification Goal | Verify that each VLSU port issues ELEN-aligned requests with `amo = AMONone`, keeps valid/payload stable until ready, and operates correctly with in-order responses of any latency, including one response per store. |
| Stimulus | Memory slave agents on every port: random `req_ready` deassertion (single cycles and long stalls), random in-order response latency 1..32 cycles, response for every write; run all load/store addressing modes on top. |
| Check | assertion + scoreboard |
| Note | SPATZ_SPEC.md: "**Responses must be returned in request order per port** …"; "**Every request, including stores, must receive exactly one response** …" |
| Old testbench situation | Real TCDM traffic, data checked end to end; TCDM has fixed latency, backpressure only from natural bank conflicts, no protocol assertions. |
| Status | In Progress |

The regression passes, yet the item is not closed: the memory timing the VLSU must tolerate is reached only in the narrow window a fixed-latency TCDM produces, and nothing checks the protocol itself.


## 3. Inputs for Phase 3 — Environment build

The following material was produced during the specification analysis and is the starting point for the UVM environment. It will move to the Phase 3 documentation once that phase is defined by the methodology.

### 3.1 Observability

| Item | Visible at the boundary | How to check |
|------|-------------------------|--------------|
| CSR values, `vl` after `vsetvl*` | `rsp_o` | direct compare |
| Scalar integer results | `rsp_o` | direct compare |
| Scalar FP results / `vfmv.f.s` | no (internal FP register file) | move to integer (`fmv.x.w`) or store (`fsw/fsd`), or backdoor `i_fpu_sequencer.i_fpr` |
| Vector register contents | no | store to memory (`vse*`) or backdoor `i_vrf.gen_reg_banks[b].gen_vrf_slice[fu].i_vregfile.mem` (bank mapping in SPATZ_SPEC.md, *Vector Register File*) |
| Memory side effects | VLSU ports, FP LSU | per-transaction monitor and/or final memory image |
| FP flags | `fpu_status_o` | OR-accumulate over a sequence and compare with the reference `fflags` |
| Instruction completion | `spatz_mem_finished_o`, `issue_rsp_o.isfloat` | end of sequence: `isfloat == 0`, no memory transaction pending, write-back queue empty |

### 3.2 Agents

| Agent | Mode | Interface | Notes |
|-------|------|-----------|-------|
| `spatz_issue_agent` | active master | `issue_valid_i/ready_o/req_i/rsp_o` | Drives instruction streams (generator or ELF/trace), samples `issue_rsp_o` in the handshake cycle; enforces A01–A04 |
| `spatz_wb_agent` | active slave | `rsp_valid_o/ready_i/rsp_o` | Random `rsp_ready_i` backpressure |
| `spatz_mem_agent` ×`NrMemPorts` | reactive slave | `spatz_mem_*[p]` | Shared memory model; random `ready` deassertion, random in-order latency, one response per write (A06, A07) |
| `spatz_fplsu_agent` | reactive slave | `fp_lsu_mem_*` | `reqrsp` slave on the same memory model (A09) |
| `spatz_side_agent` | active | `fpu_rnd_mode_i`, `fpu_fmt_mode_i` | Models `fcsr.frm`/format mode (A05); monitors `fpu_status_o` |
| `spatz_mem_finished_monitor` | passive | `spatz_mem_*finished_o` | Counts completed vector / FP memory instructions |

### 3.3 Reference model and scoreboard

- Golden model: an RVV-1.0 ISS (e.g. Spike with `--varch=vlen:512,elen:64`, as used by the software flow) or a C/SystemVerilog RVV model through DPI, configured with the Spatz behaviour of SPATZ_SPEC.md: tail-undisturbed, mask-undisturbed (except comparisons with `vma = 1`), `vill` at bit 8 of `vtype`, unsupported instructions illegal, `vxsat/vxrm/vcsr` reading 0. Q04–Q06 decide whether the model follows the RTL or RVV.
- Vector instructions produce no boundary transaction except memory traffic, so the most robust scoreboard is **end-of-sequence**: memory image, ordered list of `rsp_o` values, accumulated `fflags`. Per-store comparison is possible because stores of one instruction are issued in element order per port.
- Memory transactions of consecutive vector instructions overlap in time; load and store requests are not interleaved on the request channel, but acknowledgements of the last stores may still be in flight when the first loads of the next instruction are sent.
- Uninitialised VRF registers are X after reset (A11): the reference treats them as don't-care.

### 3.4 Functional coverage

Coverpoints, crossed where meaningful:
- instruction × SEW (8/16/32/64) × LMUL (F4, F2, 1, 2, 4, 8) × `vm` × `vl` class (0, 1, < word, = word, not multiple of word, VLMAX) × `vstart` (0 / non-zero);
- memory: addressing mode × EEW × base alignment (0..7) × stride (0, SEW, non-multiple, negative) × index EEW × `nf`;
- `vsetvl*`: AVL > / = / < VLMAX, `rs1 = x0` with `rd != x0` and `rd = x0`, keep-vl with same/different ratio, every illegal `vtype`, `vma/vta`;
- FP: rounding mode × format (incl. ALT via `fpu_fmt_mode_i`) × special operands (±0, ±inf, qNaN, sNaN, subnormal) × each flag;
- micro-architecture: in-flight instructions (1..`NrParallelInstructions`), chaining for each producer/consumer unit pair, WAR/WAW hazards, VRF bank conflicts, `BUF_FPU` FIFO full, IPU↔FPU switch, memory backpressure length, response latency, load→store and store→load switches, `rsp_ready_i` backpressure during CSR bursts;
- errors: each illegal-instruction class, instructions issued while `vill = 1`.

### 3.5 Assertions

- On every valid/ready interface: `valid && !ready |=> valid && $stable(payload)` (issue, rsp, each memory request, FP LSU).
- `spatz_mem_req_o[p].addr[2:0] == 0` and `amo == AMONone`.
- Memory side: responses per port never exceed accepted requests; at most 8 outstanding loads per port.
- `spatz_mem_str_finished_o[i] |-> spatz_mem_finished_o[i]`.
- `rsp_o.error == 0`; `rsp_o.id[5] == 0`.
- `issue_valid_i && issue_ready_o && issue_rsp_o.exception |-> !issue_rsp_o.accept`.
- After reset: `issue_rsp_o.isfloat == 0`, no memory request until the first memory instruction.
- Liveness: every accepted vector instruction retires within a bound once issue stops and memory responds.
- Internal (bind): FPU sequencer counters never roll over (`MemoryOperationCounterRollover` already in the RTL); the controller never dispatches without a free ID.
