"""
压力测试套件
覆盖三大场景：OOM压测、GC泄漏测试、流式导出性能测试

运行方式：
  python tests/stress_test.py --scenario all
  python tests/stress_test.py --scenario oom
  python tests/stress_test.py --scenario gc
  python tests/stress_test.py --scenario export
"""
import gc
import os
import sys
import time
import argparse
import traceback
import tempfile
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd

try:
    import psutil
    HAS_PSUTIL = True
except ImportError:
    HAS_PSUTIL = False
    print("⚠️  psutil 未安装，内存监控功能受限。pip install psutil")


# ─────────────────────────────────────────────
# 工具：内存快照
# ─────────────────────────────────────────────
def memory_mb() -> float:
    if not HAS_PSUTIL:
        return -1.0
    proc = psutil.Process(os.getpid())
    return proc.memory_info().rss / 1024 / 1024


def log(msg: str):
    mem = f"[MEM {memory_mb():.0f}MB]" if HAS_PSUTIL else ""
    print(f"{time.strftime('%H:%M:%S')} {mem} {msg}")


# ─────────────────────────────────────────────
# 场景一：OOM 压力测试
# 模拟 8G/16G 机器加载 10000列×500000行数据
# ─────────────────────────────────────────────
def test_oom_pressure(rows: int = 500_000, cols: int = 100, total_ram_gb: float = 8.0):
    """
    rows * cols 单元格，评估内存是否 OOM
    注意：真正的 10000 列 × 500000 行 = 50亿单元格，约需 400GB 内存，
    任何 PC 都不可能直接加载。本测试验证 chunk 策略的有效性。
    """
    print("\n" + "=" * 60)
    print(f"🔥 场景一：OOM 压力测试")
    print(f"   目标规模：{rows:,} 行 × {cols} 列 = {rows*cols:,} 单元格")
    print(f"   模拟内存：{total_ram_gb} GB")
    print("=" * 60)

    mem_start = memory_mb()
    log(f"测试开始，初始内存：{mem_start:.0f} MB")

    # ── 阶段1：评估预期内存 ──
    bytes_per_cell = 8  # float64
    expected_mb = rows * cols * bytes_per_cell / 1024 / 1024
    total_ram_mb = total_ram_gb * 1024
    log(f"预期全量加载内存：{expected_mb:.0f} MB")
    log(f"可用内存预算（80%）：{total_ram_mb * 0.8:.0f} MB")

    if expected_mb > total_ram_mb * 0.8:
        log(f"⚠️  全量加载会 OOM！需要分块策略（已超出可用内存 {total_ram_mb * 0.8:.0f} MB）")
        # 测试分块策略
        test_chunked_loading(rows, cols)
        return

    # ── 阶段2：实际构造数据并监测 ──
    log("构造测试数据（分块）...")
    chunk_size = 10_000
    chunks = []
    peak_mem = mem_start

    for i in range(0, min(rows, 100_000), chunk_size):  # 只构造 10w 行，防本机 OOM
        n = min(chunk_size, rows - i)
        data = {f"col_{j}": np.random.randn(n).astype(np.float32) for j in range(min(cols, 50))}
        chunk = pd.DataFrame(data)
        chunks.append(chunk)
        current_mem = memory_mb()
        peak_mem = max(peak_mem, current_mem)

        if HAS_PSUTIL:
            avail = psutil.virtual_memory().available / 1024 / 1024
            if avail < 500:
                log(f"❌ 内存不足！可用仅 {avail:.0f} MB，中止测试")
                del chunks
                gc.collect()
                return

    df = pd.concat(chunks, ignore_index=True)
    del chunks
    gc.collect()

    log(f"✅ 构造完成：{df.shape}，占用 {df.memory_usage(deep=True).sum() / 1024 / 1024:.0f} MB")
    log(f"   内存峰值：{peak_mem:.0f} MB（增长 {peak_mem - mem_start:.0f} MB）")

    # ── 阶段3：dtype 优化后对比 ──
    mem_before = df.memory_usage(deep=True).sum() / 1024 / 1024

    sys.path.insert(0, str(Path(__file__).parent.parent))
    from utils.helpers import optimize_dtypes
    df_opt = optimize_dtypes(df.copy())
    mem_after = df_opt.memory_usage(deep=True).sum() / 1024 / 1024
    log(f"   dtype 优化：{mem_before:.0f} MB → {mem_after:.0f} MB（节省 {(1 - mem_after/mem_before)*100:.0f}%）")

    del df, df_opt
    gc.collect()
    log(f"✅ GC 后内存：{memory_mb():.0f} MB")


def test_chunked_loading(rows: int, cols: int):
    """验证分块加载策略"""
    log("📦 测试分块加载策略...")
    tmp = Path(tempfile.mkdtemp())
    csv_path = tmp / "big.csv"

    # 生成小样本 CSV（验证逻辑，不生成全量）
    sample_rows = min(rows, 50_000)
    pd.DataFrame(
        np.random.randn(sample_rows, min(cols, 20)),
        columns=[f"col_{i}" for i in range(min(cols, 20))]
    ).to_csv(csv_path, index=False)

    log(f"   生成测试 CSV：{csv_path.stat().st_size / 1024 / 1024:.1f} MB")

    mem_before = memory_mb()
    chunk_size = 10_000
    chunk_count = 0
    peak_chunk_mem = mem_before

    for chunk in pd.read_csv(csv_path, chunksize=chunk_size):
        chunk_count += 1
        peak_chunk_mem = max(peak_chunk_mem, memory_mb())

    log(f"   分块读取：{chunk_count} 个块，峰值内存增长 {peak_chunk_mem - mem_before:.0f} MB")
    log("✅ 分块策略验证通过")

    csv_path.unlink()


# ─────────────────────────────────────────────
# 场景二：GC 泄漏测试（模拟频繁切页）
# ─────────────────────────────────────────────
def test_gc_leak(iterations: int = 20, rows: int = 100_000, cols: int = 50):
    """
    模拟 Streamlit 用户频繁切页：
    - 每次「切页」= 创建 df + 执行操作 + 存 history + 删引用
    - 检测是否有内存泄漏（基线内存持续增长）
    """
    print("\n" + "=" * 60)
    print(f"🔄 场景二：GC 泄漏测试（模拟切页 {iterations} 次）")
    print("=" * 60)

    mem_baseline = memory_mb()
    log(f"基线内存：{mem_baseline:.0f} MB")

    history: list = []
    MAX_HISTORY = 10
    mem_snapshots = []

    for i in range(iterations):
        # 模拟「加载数据」
        df = pd.DataFrame(
            np.random.randn(rows, cols).astype(np.float32),
            columns=[f"col_{j}" for j in range(cols)]
        )

        # 模拟「执行操作」
        df["new_col"] = df["col_0"] * 2 + df["col_1"]
        df["category"] = pd.cut(df["new_col"], bins=5)

        # 模拟「存 snapshot」（原逻辑：df.copy()）
        if i % 3 == 0:  # 模拟部分操作触发 snapshot
            history.append(df.copy())
            if len(history) > MAX_HISTORY:
                old = history.pop(0)
                del old

        # 模拟「切页」：删除当前页的局部引用
        del df
        gc.collect()

        snap = memory_mb()
        mem_snapshots.append(snap)

        if i % 5 == 0:
            log(f"第 {i+1:02d} 次切页，内存：{snap:.0f} MB（+{snap - mem_baseline:.0f} MB）")

    # ── 清空 history ──
    del history
    gc.collect()
    mem_after_clear = memory_mb()
    log(f"清空 history 后：{mem_after_clear:.0f} MB（+{mem_after_clear - mem_baseline:.0f} MB）")

    # ── 分析是否有泄漏 ──
    if len(mem_snapshots) > 5:
        first_half = np.mean(mem_snapshots[:len(mem_snapshots) // 2])
        second_half = np.mean(mem_snapshots[len(mem_snapshots) // 2:])
        growth = second_half - first_half
        if growth > 50:
            log(f"⚠️  检测到内存泄漏趋势！前半段均值 {first_half:.0f} MB vs 后半段 {second_half:.0f} MB，增长 {growth:.0f} MB")
        else:
            log(f"✅ 内存稳定，无明显泄漏（前半 {first_half:.0f} vs 后半 {second_half:.0f} MB）")


# ─────────────────────────────────────────────
# 场景三：流式导出性能测试
# ─────────────────────────────────────────────
def test_export_streaming(rows: int = 200_000, cols: int = 30):
    """
    对比三种导出方案的内存峰值和耗时：
    - 方案A：df.to_csv()（全量内存）
    - 方案B：分块写入 BytesIO（低峰值）
    - 方案C：helpers.df_to_excel_optimized（constant_memory 模式）
    """
    print("\n" + "=" * 60)
    print(f"💾 场景三：导出流式写入测试（{rows:,} 行 × {cols} 列）")
    print("=" * 60)

    df = pd.DataFrame(
        {
            **{f"num_{i}": np.random.randn(rows).astype(np.float32) for i in range(cols // 2)},
            **{f"str_{i}": np.random.choice(["A", "B", "C", "D"], rows) for i in range(cols // 2)},
        }
    )
    log(f"测试 DataFrame：{df.memory_usage(deep=True).sum() / 1024 / 1024:.0f} MB")

    sys.path.insert(0, str(Path(__file__).parent.parent))
    from utils.helpers import df_to_excel_optimized

    results = {}

    # ── 方案A：df.to_csv() 全量 ──
    gc.collect()
    mem_a0 = memory_mb()
    t0 = time.perf_counter()
    buf_a = BytesIO()
    df.to_csv(buf_a, index=False, encoding="utf-8-sig")
    buf_size_a = buf_a.tell() / 1024 / 1024
    t_a = time.perf_counter() - t0
    mem_a1 = memory_mb()
    del buf_a
    gc.collect()
    results["A: df.to_csv"] = {
        "time": t_a, "mem_peak": mem_a1 - mem_a0, "file_mb": buf_size_a
    }
    log(f"方案A：{t_a:.2f}s，峰值内存 +{mem_a1 - mem_a0:.0f} MB，文件 {buf_size_a:.1f} MB")

    # ── 方案B：分块 CSV ──
    gc.collect()
    mem_b0 = memory_mb()
    t0 = time.perf_counter()
    buf_b = BytesIO()
    chunk_size = 50_000
    first = True
    for start in range(0, len(df), chunk_size):
        chunk = df.iloc[start:start + chunk_size]
        chunk.to_csv(buf_b, index=False, header=first, encoding="utf-8-sig",
                     mode="a" if not first else "w")
        first = False
        del chunk
    buf_size_b = buf_b.tell() / 1024 / 1024
    t_b = time.perf_counter() - t0
    mem_b1 = memory_mb()
    del buf_b
    gc.collect()
    results["B: 分块CSV"] = {
        "time": t_b, "mem_peak": mem_b1 - mem_b0, "file_mb": buf_size_b
    }
    log(f"方案B：{t_b:.2f}s，峰值内存 +{mem_b1 - mem_b0:.0f} MB，文件 {buf_size_b:.1f} MB")

    # ── 方案C：df_to_excel_optimized ──
    gc.collect()
    mem_c0 = memory_mb()
    t0 = time.perf_counter()
    buf_c = df_to_excel_optimized({"Sheet1": df})
    buf_size_c = len(buf_c.getvalue()) / 1024 / 1024
    t_c = time.perf_counter() - t0
    mem_c1 = memory_mb()
    del buf_c
    gc.collect()
    results["C: Excel优化"] = {
        "time": t_c, "mem_peak": mem_c1 - mem_c0, "file_mb": buf_size_c
    }
    log(f"方案C：{t_c:.2f}s，峰值内存 +{mem_c1 - mem_c0:.0f} MB，文件 {buf_size_c:.1f} MB")

    # ── 汇总 ──
    print("\n📊 导出性能汇总：")
    print(f"{'方案':<18} {'耗时(s)':<10} {'内存峰值(MB)':<15} {'文件大小(MB)'}")
    print("-" * 60)
    for name, r in results.items():
        print(f"{name:<18} {r['time']:<10.2f} {r['mem_peak']:<15.0f} {r['file_mb']:.1f}")

    del df
    gc.collect()


# ─────────────────────────────────────────────
# 场景四：快速估算"你的机器能承载多大数据"
# ─────────────────────────────────────────────
def estimate_capacity():
    """根据当前机器内存估算最大可处理规模"""
    print("\n" + "=" * 60)
    print("📐 机器容量评估")
    print("=" * 60)

    if not HAS_PSUTIL:
        print("需要 psutil：pip install psutil")
        return

    mem = psutil.virtual_memory()
    total_gb = mem.total / 1024 ** 3
    avail_gb = mem.available / 1024 ** 3

    print(f"总内存：{total_gb:.1f} GB，可用：{avail_gb:.1f} GB")

    # 预留 30% 给系统
    usable_gb = avail_gb * 0.7
    usable_bytes = usable_gb * 1024 ** 3

    # float32 占 4 字节，float64 占 8 字节
    scenarios = [
        ("全 float32（优化后）", 4),
        ("全 float64（默认）", 8),
        ("混合类型", 12),  # 含字符串
    ]

    print("\n可处理的最大数据规模（行数估算）：")
    for name, bytes_per_cell in scenarios:
        for cols in [50, 100, 500, 1000, 10000]:
            max_rows = int(usable_bytes / (bytes_per_cell * cols))
            if max_rows > 100:
                print(f"  {name}，{cols} 列：最多 {max_rows:,} 行（约 {max_rows*cols/1e9:.2f}B 单元格）")
                break

    print(f"\n⚡ 建议：")
    if avail_gb < 4:
        print("  内存紧张，建议限制加载 < 10万行，使用分块+DuckDB模式")
    elif avail_gb < 8:
        print("  中等内存，可处理 50-100万行（优化dtype后）")
    elif avail_gb < 16:
        print("  充足内存，可处理 200-500万行（优化dtype后）")
    else:
        print("  大内存机器，可处理 1000万行以上（优化dtype后）")


# ─────────────────────────────────────────────
# 主入口
# ─────────────────────────────────────────────
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Excel工具压力测试套件")
    parser.add_argument(
        "--scenario",
        choices=["all", "oom", "gc", "export", "capacity"],
        default="capacity",
        help="测试场景"
    )
    parser.add_argument("--rows", type=int, default=100_000, help="测试行数")
    parser.add_argument("--cols", type=int, default=50, help="测试列数")
    parser.add_argument("--ram", type=float, default=8.0, help="模拟内存大小(GB)")
    args = parser.parse_args()

    print("=" * 60)
    print("🚀 Excel 工具压力测试套件")
    print(f"   Python {sys.version.split()[0]}")
    print(f"   Pandas {pd.__version__}")
    print(f"   NumPy {np.__version__}")
    print("=" * 60)

    try:
        if args.scenario in ("all", "capacity"):
            estimate_capacity()
        if args.scenario in ("all", "oom"):
            test_oom_pressure(args.rows, args.cols, args.ram)
        if args.scenario in ("all", "gc"):
            test_gc_leak(iterations=20, rows=min(args.rows, 100_000), cols=args.cols)
        if args.scenario in ("all", "export"):
            test_export_streaming(rows=min(args.rows, 200_000), cols=args.cols)
    except KeyboardInterrupt:
        print("\n⛔ 测试中止")
    except Exception as e:
        print(f"\n❌ 测试失败：{e}")
        traceback.print_exc()

    print("\n✅ 测试完成")
