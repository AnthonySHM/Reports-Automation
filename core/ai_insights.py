"""
AI-Powered Cyber Insight Generator

Generates contextual observations and recommendations for Slide 3
based on the report's security metrics and findings.
"""

from __future__ import annotations

import logging
import os
from typing import Optional

logger = logging.getLogger("shm.ai_insights")


def generate_cyber_insight(
    microsoft_kb_count: int,
    software_count: int,
    key_detections: int,
    soc_tickets: int,
    analyst_investigations: int,
    nuclei_vulns_count: int = 0,
) -> str:
    """Generate an AI-powered observation for the Cyber Insight slide.
    
    Creates a concise, professional observation based on the report metrics.
    
    Parameters
    ----------
    microsoft_kb_count : int
        Number of missing Microsoft KB patches
    software_count : int
        Number of missing software packages
    key_detections : int
        Number of key security detections
    soc_tickets : int
        Number of SOC tickets generated
    analyst_investigations : int
        Number of analyst investigations
    nuclei_vulns_count : int
        Number of network vulnerabilities found
    
    Returns
    -------
    str
        A single-sentence observation/recommendation
    """
    insights = []
    
    # Patch-related insights
    if microsoft_kb_count > 0 or software_count > 0:
        total_patches = microsoft_kb_count + software_count
        if total_patches > 10:
            insights.append(
                "The environment shows a moderate patching backlog that should be addressed "
                "systematically to reduce the attack surface."
            )
        elif total_patches > 5:
            insights.append(
                "The organization maintains a manageable patch status that should be addressed "
                "in the next maintenance window."
            )
        elif total_patches > 0:
            insights.append(
                "The environment demonstrates strong patch management with minimal pending updates remaining."
            )
    
    # SOC activity insights
    if soc_tickets > 10:
        insights.append(
            "Elevated SOC activity indicates heightened monitoring vigilance during this period, "
            "with most events requiring client confirmation before escalation."
        )
    elif soc_tickets > 5:
        insights.append(
            "Normal SOC activity levels reflect routine security monitoring and responsive threat detection."
        )
    elif soc_tickets > 0:
        insights.append(
            "Minimal SOC escalations indicate a stable security posture with few anomalous events "
            "requiring client attention."
        )
    
    # Key detections insights
    if key_detections > 100:
        insights.append(
            "High-severity detection volume suggests active threat landscape requiring continued "
            "monitoring and response capabilities."
        )
    elif key_detections > 50:
        insights.append(
            "Moderate detection activity is consistent with normal enterprise security operations."
        )
    
    # Network vulnerability insights
    if nuclei_vulns_count > 10:
        insights.append(
            "Internal network scanning identified multiple vulnerabilities, recommending prioritized "
            "remediation based on severity and exploitability."
        )
    elif nuclei_vulns_count > 0:
        insights.append(
            "Network scanning revealed potential vulnerabilities that should be reviewed and addressed "
            "in accordance with risk tolerance."
        )
    
    # Analyst investigation insights
    if analyst_investigations > 20:
        insights.append(
            "Extensive analyst review reflects thorough examination of contextual indicators and "
            "proactive threat hunting."
        )
    
    # Default if no specific insights
    if not insights:
        return (
            "The reporting period reflects stable security operations with routine "
            "monitoring activities and no significant anomalies requiring immediate action."
        )
    
    # Return the first (most relevant) insight
    return insights[0]


def generate_simple_insight(
    microsoft_kb_count: int,
    software_count: int,
) -> str:
    """Generate a simple insight based only on patch counts.
    
    Fallback when other metrics aren't available.
    """
    total = microsoft_kb_count + software_count
    
    if total > 15:
        return (
            "The organization should prioritize patch deployment for pending Microsoft KB "
            "updates and software packages to maintain a strong security posture."
        )
    elif total > 8:
        return (
            "Current patch requirements are within acceptable levels and should be "
            "addressed during the next scheduled maintenance window."
        )
    elif total > 0:
        return (
            "The environment demonstrates excellent patch management with minimal "
            "outstanding updates, reflecting proactive vulnerability management practices."
        )
    else:
        return (
            "The environment maintains up-to-date patch status with no pending critical "
            "updates identified during the reporting period."
        )

