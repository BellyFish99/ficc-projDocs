# Solution Skeleton Patterns

Standardized architectural components and strategic frameworks for FinTech, FICC, and POMS solutions.

## 1. The "Capability Leap" Framework (5D)
Use this framework to organize any solution skeleton for a large trading/investment client.
- **[配得优] Allocation**: Multi-asset, unified portfolio model (Position Model).
- **[算得快] Calculation**: Real-time pricing, Greeks, VaR, and all-in cost engines.
- **[控得稳] Control**: Pre-trade compliance, liquidity stress testing, and limit management.
- **[连得通] Connection**: STP (Straight Through Processing), OEMS integration, and research signals.
- **[领得先] Leading**: Localization (信创), cloud-native, and future-proofing.

## 2. POMS Engine Architecture
A standard POMS must contain these 6 engines:
1. **Portfolio Management**: Unified views, what-if simulation, and rebalancing.
2. **Calculation Engine**: Pricing models (MC, B-S), valuation, and risk sensitivities.
3. **Strategy Engine**: RV (Relative Value), arbitrage, and multi-timeframe signals.
4. **Risk & Compliance**: CEP (Complex Event Processing), pre-trade check, and drawdown alerts.
5. **Instruction & Execution**: OMS (Order Management), multi-leg execution, and TCA.
6. **Accounting & Cost**: IFRS9, financing costs (Repo), and CFETS/Settlement fee attribution.

## 3. Best-in-Class Technical Blueprint (Industry Standard)
When designing for scale and latency (e.g., ¥100B+ AUM), adopt these technical components:
- **Event-Driven Backbone**: High-performance messaging (Aeron for low-latency, Kafka for high-throughput) for real-time trade data propagation.
- **Microservices Architecture**: Decoupled services for Portfolio (PM), Order Routing (SOR), and Risk for independent scaling.
- **High-Concurrency Data Fabric**: Unified source of truth (Golden Source) for positions and market data using distributed databases (e.g., TiDB).
- **Embedded Compliance**: Millisecond-level pre-trade checks integrated directly into the order life cycle.
- **IFRS9 Accounting Engine**: Automated FVTPL/FVOCI/AC classification and real-time ECL (Expected Credit Loss) modeling.

## 4. The "Automated Ecosystem" Metaphor
When integrating with existing systems (IBOR, Market Data):
- **Brain**: POMS (Logic, Intelligence, Decision).
- **Engine**: IBOR (Data, Bookkeeping).
- **Sensors**: Market Data/Feeds (Real-time updates).
- **Chassis/Body**: OEMS/Counter systems (Execution).

## 5. ROI Value Formulas (Conservative Estimates)
Include these in the "Value" section of the skeleton:
- **Alpha Improvement**: `AUM * 0.02% (2bps) = ¥X/year`.
- **Slippage Reduction**: `Annual Trading Volume * 0.5bp = ¥Y/year`.
- **Financing Optimization**: `Repo Size * 5bps = ¥Z/year`.
- **Risk Avoidance**: `One-time Extreme Loss * Probability = ¥W`.
