"""
Patch History Management

Tracks historical patch counts over time and generates trend charts
that automatically update with the current reporting period's data.
"""

from __future__ import annotations

import csv
import json
import logging
from datetime import datetime
from pathlib import Path
from typing import Optional

logger = logging.getLogger("shm.patch_history")


class PatchHistoryManager:
    """Manages historical patch count data for trend chart generation."""
    
    def __init__(self, storage_path: str | Path = "assets/patch_history"):
        """Initialize the patch history manager.
        
        Parameters
        ----------
        storage_path : str | Path
            Directory where patch history JSON files are stored
        """
        self.storage_path = Path(storage_path)
        self.storage_path.mkdir(parents=True, exist_ok=True)
    
    def _get_history_file(self, client_slug: str) -> Path:
        """Get the path to a client's history file."""
        return self.storage_path / f"{client_slug}_patch_history.json"
    
    def load_history(self, client_slug: str) -> list[dict]:
        """Load patch history for a client.
        
        Returns
        -------
        list[dict]
            List of history entries, each with: date, microsoft_count, software_count
        """
        history_file = self._get_history_file(client_slug)
        
        if not history_file.exists():
            logger.info("No history file found for '%s', starting fresh", client_slug)
            return []
        
        try:
            with open(history_file, 'r', encoding='utf-8') as f:
                history = json.load(f)
            logger.info("Loaded %d history entries for '%s'", len(history), client_slug)
            return history
        except Exception as exc:
            logger.error("Failed to load history for '%s': %s", client_slug, exc)
            return []
    
    def save_history(self, client_slug: str, history: list[dict]) -> None:
        """Save patch history for a client.
        
        Parameters
        ----------
        client_slug : str
            Client identifier
        history : list[dict]
            List of history entries
        """
        history_file = self._get_history_file(client_slug)
        
        try:
            with open(history_file, 'w', encoding='utf-8') as f:
                json.dump(history, f, indent=2)
            logger.info("Saved %d history entries for '%s'", len(history), client_slug)
        except Exception as exc:
            logger.error("Failed to save history for '%s': %s", client_slug, exc)
    
    def add_entry(
        self,
        client_slug: str,
        end_date: str,
        microsoft_count: int,
        software_count: int,
        max_history: int = 6
    ) -> None:
        """Add or update a history entry for the current report period.
        
        Parameters
        ----------
        client_slug : str
            Client identifier
        end_date : str
            Report period end date (e.g., "2026-01-05")
        microsoft_count : int
            Count of missing Microsoft KB patches
        software_count : int
            Count of missing software packages
        max_history : int
            Maximum number of history entries to keep (default: 6)
        """
        history = self.load_history(client_slug)
        
        # Check if entry for this date already exists
        existing_idx = None
        for i, entry in enumerate(history):
            if entry.get('date') == end_date:
                existing_idx = i
                break
        
        new_entry = {
            'date': end_date,
            'microsoft_count': microsoft_count,
            'software_count': software_count,
            'updated_at': datetime.now().isoformat(),
        }
        
        if existing_idx is not None:
            # Update existing entry
            history[existing_idx] = new_entry
            logger.info("Updated history entry for '%s' on %s", client_slug, end_date)
        else:
            # Add new entry
            history.append(new_entry)
            logger.info("Added history entry for '%s' on %s", client_slug, end_date)
        
        # Sort by date
        history.sort(key=lambda x: x['date'])
        
        # Keep only the most recent entries
        if len(history) > max_history:
            history = history[-max_history:]
            logger.info("Trimmed history to %d most recent entries", max_history)
        
        self.save_history(client_slug, history)
    
    def get_trend_data(
        self,
        client_slug: str,
        end_date: str,
        microsoft_count: int,
        software_count: int,
        num_points: int = 3
    ) -> tuple[list[str], list[int], list[int]]:
        """Get trend data for chart generation.
        
        Automatically includes the current report period as the last data point,
        updating or adding it to the historical data.
        
        Parameters
        ----------
        client_slug : str
            Client identifier
        end_date : str
            Current report period end date (e.g., "2026-01-05")
        microsoft_count : int
            Current count of missing Microsoft KB patches
        software_count : int
            Current count of missing software packages
        num_points : int
            Number of data points to include in the trend (default: 3)
        
        Returns
        -------
        tuple[list[str], list[int], list[int]]
            (dates, microsoft_counts, software_counts) for chart generation
            Dates are formatted as "Month Day" (e.g., "January 5")
        """
        # Add/update current entry
        self.add_entry(client_slug, end_date, microsoft_count, software_count)
        
        # Load updated history
        history = self.load_history(client_slug)
        
        if not history:
            logger.warning("No history available for '%s'", client_slug)
            return [], [], []
        
        # Take the most recent N entries
        recent = history[-num_points:]
        
        # Format dates for display
        dates = []
        for entry in recent:
            try:
                date_obj = datetime.strptime(entry['date'], '%Y-%m-%d')
                # Format as "Month Day" (e.g., "January 5")
                formatted = date_obj.strftime('%B %d')
                # Remove leading zero from day (e.g., "January 05" -> "January 5")
                parts = formatted.split()
                if len(parts) == 2 and parts[1].startswith('0'):
                    formatted = f"{parts[0]} {parts[1].lstrip('0')}"
                dates.append(formatted)
            except:
                # Fallback: use raw date
                dates.append(entry['date'])
        
        microsoft_counts = [entry['microsoft_count'] for entry in recent]
        software_counts = [entry['software_count'] for entry in recent]
        
        logger.info(
            "Generated trend data for '%s': %d points, latest = MS:%d SW:%d",
            client_slug, len(dates), microsoft_counts[-1], software_counts[-1]
        )
        
        return dates, microsoft_counts, software_counts


def generate_patch_trend_chart_for_report(
    client_slug: str,
    end_date: str,
    microsoft_count: int,
    software_count: int,
    output_path: str | Path,
    *,
    num_points: int = 3,
    storage_path: str | Path = "assets/patch_history",
) -> Path:
    """Generate a patch trend chart for a report, automatically managing history.
    
    This is the main function to call from the API. It:
    1. Loads historical patch data for the client
    2. Adds/updates the current report period's counts
    3. Generates a trend chart with the most recent data points
    4. Saves the updated history for future reports
    
    Parameters
    ----------
    client_slug : str
        Client identifier (e.g., "elephant", "wiegmann")
    end_date : str
        Report period end date (format: "YYYY-MM-DD", e.g., "2026-01-05")
    microsoft_count : int
        Current count of missing Microsoft KB patches
    software_count : int
        Current count of missing software packages
    output_path : str | Path
        Where to save the generated chart image
    num_points : int
        Number of historical data points to show (default: 3)
    storage_path : str | Path
        Directory for patch history storage (default: "assets/patch_history")
    
    Returns
    -------
    Path
        Path to the generated chart image
    
    Examples
    --------
    >>> chart_path = generate_patch_trend_chart_for_report(
    ...     client_slug="elephant",
    ...     end_date="2026-01-05",
    ...     microsoft_count=4,
    ...     software_count=2,
    ...     output_path="assets/ndr_cache/elephant/patch_trend.png"
    ... )
    """
    from core.patch_chart import generate_patch_trend_chart
    
    manager = PatchHistoryManager(storage_path)
    
    # Get trend data (automatically updates history)
    dates, ms_counts, sw_counts = manager.get_trend_data(
        client_slug=client_slug,
        end_date=end_date,
        microsoft_count=microsoft_count,
        software_count=software_count,
        num_points=num_points,
    )
    
    if not dates:
        logger.warning(
            "No trend data available for '%s', cannot generate chart",
            client_slug
        )
        return None
    
    # Generate the chart
    return generate_patch_trend_chart(
        dates=dates,
        microsoft_counts=ms_counts,
        software_counts=sw_counts,
        output_path=output_path,
    )


# Convenience function for testing
if __name__ == "__main__":
    import time
    
    # Test the history manager
    manager = PatchHistoryManager("assets/patch_history")
    
    # Simulate adding historical entries
    test_data = [
        ("2025-11-10", 6, 3),
        ("2025-12-08", 5, 3),
        ("2026-01-05", 4, 2),
    ]
    
    for date, ms, sw in test_data:
        manager.add_entry("test_client", date, ms, sw)
    
    # Get trend data
    dates, ms_counts, sw_counts = manager.get_trend_data(
        "test_client", "2026-01-05", 4, 2
    )
    
    print(f"Dates: {dates}")
    print(f"Microsoft counts: {ms_counts}")
    print(f"Software counts: {sw_counts}")
