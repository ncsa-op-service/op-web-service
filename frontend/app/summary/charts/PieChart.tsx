import "./PieChart.css";

type ChartItem = {
  label: string;
  value: number;
};

type PieChartProps = {
  data: ChartItem[];
};

export default function PieChart({
  data,
}: PieChartProps) {
  const total = data.reduce(
    (sum, item) =>
      sum + item.value,
    0
  );

  const safeTotal =
    total > 0 ? total : 1;

  let currentPercentage = 0;

  const segments = data.map(
    (item, index) => {
      const start =
        currentPercentage;

      const percentage =
        (item.value / safeTotal) *
        100;

      currentPercentage += percentage;

      return `var(--pie-color-${index}) ${start}% ${currentPercentage}%`;
    }
  );

  return (
    <div className="pie-chart-container">
      <div
        className="pie-chart"
        style={{
          background:
            total > 0
              ? `conic-gradient(${segments.join(
                  ", "
                )})`
              : "#e7ecf3",
        }}
      >
        <div className="pie-hole">
          <strong>{total}</strong>
          <span>ทั้งหมด</span>
        </div>
      </div>

      <div className="pie-legend">
        {data.map(
          (item, index) => {
            const percent =
              total > 0
                ? Math.round(
                    (item.value /
                      total) *
                      100
                  )
                : 0;

            return (
              <div
                className="pie-legend-item"
                key={item.label}
              >
                <span
                  className={`pie-dot pie-dot-${index}`}
                />

                <div>
                  <strong>
                    {item.label}
                  </strong>

                  <p>
                    {item.value} รายการ ·{" "}
                    {percent}%
                  </p>
                </div>
              </div>
            );
          }
        )}
      </div>
    </div>
  );
}