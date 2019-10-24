# CSA-DistSyst
A novel N-1 contingency screening approach for distribution system (Academic use only)

This is a repository of data and code for PES General Meeting 2020 paper "A Novel Acceleration Strategy for N-1 Contingency Screening in Distribution System". 

# Abstract
The concern for the security of distribution network renders the necessity of developing an advanced N-1 contingency screening method. However, the complexity of network reconfiguration and islanding operation brings tremendous computation burden for the traditional screening approach. To cope with the problem, a novel acceleration strategy is proposed based on an idea of integrating the optimal power flow runs under every N-1 contingency into a single search tree. To dynamically check the N-1 security criterion during the solving procedure, the model is formulated in a branch-and-cut framework. At each incumbent solution node, a specialized lazy constraint callback approach is implementedto eliminate the current N-1 contingency from the contingency set if the load shedding can be resolved by network reconfiguration or islanding operation. Thus, all calculations for the N-1 contingency are incorporated as a whole and performed only once. The superiority and efficiency of the strategy are varified on the IEEE-33 and 69 bus systems. Numerical results indicate that the proposed approach is an order of magnitude faster than the traditional ones.

# Research Interests
We are looking forward to cooperate with all experts and scholars around the world on the topic of Big Data and Machine Learning in power system.
