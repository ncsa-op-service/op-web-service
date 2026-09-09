import "./LineChart.css";

type ChartItem = {
  label: string;
  value: number;
};

type LineChartProps = {
  data: ChartItem[];
};

export default function LineChart({
  data,
}: LineChartProps) {
  const width = 1000;
  const height = 260;

  const paddingLeft = 70;
  const paddingRight = 50;
  const paddingTop = 35;
  const paddingBottom = 55;

  const maxValue = Math.max(
    ...data.map((item) => item.value),
    1
  );

  const graphWidth =
    width -
    paddingLeft -
    paddingRight;

  const graphHeight =
    height -
    paddingTop -
    paddingBottom;

  const points = data.map(
    (item, index) => {
      const x =
        data.length <= 1
          ? width / 2
          : paddingLeft +
            (index /
              (data.length - 1)) *
              graphWidth;

      const y =
        paddingTop +
        graphHeight -
        (item.value /
          maxValue) *
          graphHeight;

      return {
        ...item,
        x,
        y,
      };
    }
  );

  const pointString =
    points
      .map(
        (point) =>
          `${point.x},${point.y}`
      )
      .join(" ");

  const steps = 4;

  const gridValues =
    Array.from(
      {
        length: steps + 1,
      },
      (_, index) =>
        Math.round(
          (maxValue / steps) *
            index
        )
    ).reverse();

  return (
    <div className="line-chart-container">
      <svg
        className="line-chart-svg"
        viewBox={`0 0 ${width} ${height}`}
        preserveAspectRatio="none"
      >
        {gridValues.map(
          (value, index) => {
            const y =
              paddingTop +
              (index / steps) *
                graphHeight;

            return (
              <g key={value}>
                <line
                  x1={paddingLeft}
                  x2={
                    width -
                    paddingRight
                  }
                  y1={y}
                  y2={y}
                  className="line-grid"
                />

                <text
                  x={
                    paddingLeft -
                    15
                  }
                  y={y + 4}
                  textAnchor="end"
                  className="line-y-label"
                >
                  {value}
                </text>
              </g>
            );
          }
        )}

        <polyline
          points={pointString}
          fill="none"
          className="line-path"
        />

        {points.map(
          (point) => (
            <g key={point.label}>
              <circle
                cx={point.x}
                cy={point.y}
                r="6"
                className="line-point"
              />

              <text
                x={point.x}
                y={point.y - 14}
                textAnchor="middle"
                className="line-value"
              >
                {point.value}
              </text>

              <text
                x={point.x}
                y={
                  height -
                  18
                }
                textAnchor="middle"
                className="line-x-label"
              >
                {point.label}
              </text>
            </g>
          )
        )}
      </svg>
    </div>
  );
}