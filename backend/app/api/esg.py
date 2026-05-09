import json
import logging
import os
from collections import defaultdict
from typing import Any, Dict, List, Optional

import yaml
from fastapi import APIRouter
from pydantic import BaseModel

from app.services.llm_client import llm_json

router = APIRouter(prefix="/api/v1/esg", tags=["esg"])
logger = logging.getLogger(__name__)

ANALYSIS_PROMPT_PATH = os.path.join(os.path.dirname(__file__), "..", "prompts", "v1", "esg_analysis.yaml")
REPORT_PROMPT_PATH = os.path.join(os.path.dirname(__file__), "..", "prompts", "v1", "esg_report.yaml")


def _load_prompt(path: str) -> dict:
    with open(path, "r") as f:
        return yaml.safe_load(f)


class AssetInput(BaseModel):
    id: str
    name: str
    category: str
    super_category: Optional[str] = None
    province: str
    autonomous_community: str
    municipality: str
    latitude: float
    longitude: float
    address: str
    confidence_score: float
    data_sources: List[str] = []


class EsgAnalyzeRequest(BaseModel):
    company_name: str
    assets: List[AssetInput]


class EsgReportRequest(BaseModel):
    company_name: str
    analysis: Dict[str, Any]


_MAX_ESG_ASSETS = 40


def _sample_representative(assets: List[AssetInput], limit: int) -> List[AssetInput]:
    """
    Keep up to `limit` assets while preserving diversity across categories and provinces.
    Always keeps the highest-confidence asset per (category, province) bucket first,
    then fills remaining slots with the top-confidence remainder.
    """
    if len(assets) <= limit:
        return assets

    # Sort by confidence descending so best assets are picked first
    sorted_assets = sorted(assets, key=lambda a: a.confidence_score, reverse=True)

    # Pick one representative per (category, province) bucket
    seen_buckets: set = set()
    primary: List[AssetInput] = []
    rest: List[AssetInput] = []
    for a in sorted_assets:
        bucket = (a.category, a.province)
        if bucket not in seen_buckets:
            seen_buckets.add(bucket)
            primary.append(a)
        else:
            rest.append(a)

    selected = primary[:limit]
    if len(selected) < limit:
        selected += rest[: limit - len(selected)]

    return selected


@router.post("/analyze")
async def analyze_esg(req: EsgAnalyzeRequest):
    prompt = _load_prompt(ANALYSIS_PROMPT_PATH)

    total_assets = len(req.assets)
    sampled = _sample_representative(req.assets, _MAX_ESG_ASSETS)
    was_sampled = len(sampled) < total_assets

    assets_data = [
        {
            "id": a.id,
            "name": a.name,
            "category": a.category,
            "super_category": a.super_category or "OTR",
            "province": a.province,
            "autonomous_community": a.autonomous_community,
            "municipality": a.municipality,
        }
        for a in sampled
    ]

    sampling_note = (
        f" (muestra representativa de {len(sampled)} activos sobre un total de {total_assets})"
        if was_sampled
        else ""
    )

    system_msg = prompt["system"].replace("{company_name}", req.company_name)
    user_msg = (
        prompt["user"]
        .replace("{company_name}", req.company_name)
        .replace("{assets_json}", json.dumps(assets_data, ensure_ascii=False, indent=2))
        .replace("{asset_count}", str(total_assets) + sampling_note)
    )

    result = await llm_json(
        [
            {"role": "system", "content": system_msg},
            {"role": "user", "content": user_msg},
        ],
        temperature=0.2,
        timeout=180,
        max_tokens=8192,
    )

    if not result:
        return {"error": "LLM analysis failed"}

    return result


@router.post("/report")
async def generate_report(req: EsgReportRequest):
    prompt = _load_prompt(REPORT_PROMPT_PATH)

    system_msg = prompt["system"].replace("{company_name}", req.company_name)
    user_msg = (
        prompt["user"]
        .replace("{company_name}", req.company_name)
        .replace("{analysis_json}", json.dumps(req.analysis, ensure_ascii=False, indent=2))
    )

    result = await llm_json(
        [
            {"role": "system", "content": system_msg},
            {"role": "user", "content": user_msg},
        ],
        temperature=0.3,
        timeout=180,
    )

    if not result:
        return {"error": "LLM report generation failed"}

    return result
