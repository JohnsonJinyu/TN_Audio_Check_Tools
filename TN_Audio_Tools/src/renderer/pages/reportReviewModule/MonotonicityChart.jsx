import React from 'react';
import { LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, ReferenceDot, ResponsiveContainer } from 'recharts';

const DIR_LABELS = {
  RCV: '接收方向 (RCV)',
  SND: '发送方向 (SND)',
  TX: '发送方向 (TX)',
  RX: '接收方向 (RX)',
};

/**
 * 检查数据是否单调
 * 返回 { monotonic: boolean, trend: 'up'|'down'|'flat', violations: [{index, level, value}] }
 */
function analyzeMonotonicity(points) {
  if (points.length < 2) return { monotonic: true, trend: 'flat', violations: [] };
  var ups = 0, downs = 0;
  for (var i = 1; i < points.length; i++) {
    if (points[i].value > points[i - 1].value) ups++;
    else if (points[i].value < points[i - 1].value) downs++;
  }
  var trend = ups > downs ? 'up' : downs > ups ? 'down' : 'flat';
  var violations = [];
  for (var i = 1; i < points.length; i++) {
    var diff = points[i].value - points[i - 1].value;
    var isOk = trend === 'up' ? diff >= -0.3 : trend === 'down' ? diff <= 0.3 : Math.abs(diff) < 0.5;
    if (!isOk) {
      violations.push({ index: i, level: points[i].level, value: points[i].value, diff: diff });
    }
  }
  return { monotonic: violations.length === 0, trend: trend, violations: violations };
}

export default function MonotonicityChart({ data }) {
  if (!data || Object.keys(data).length === 0) return null;

  return (
    <div style={{ marginTop: 16 }}>
      <h4 style={{ marginBottom: 8 }}>Loudness Rating 值-等级趋势图：</h4>
      <div style={{ marginBottom: 12, fontSize: 12, color: 'var(--text-light)' }}>
        本图展示的是 RLR / SLR rating 数值，不是实际声压级；对 loudness rating 而言，dB 值越小通常表示实际越响，因此图中上方代表更响、下方代表更弱。
      </div>
      {Object.keys(data).map(function(dir) {
        var points = data[dir];
        if (!points || points.length < 2) return null;
        var analysis = analyzeMonotonicity(points);
        var minVal = Math.min.apply(null, points.map(function(p) { return p.value; }));
        var maxVal = Math.max.apply(null, points.map(function(p) { return p.value; }));
        var padding = Math.max(1, (maxVal - minVal) * 0.2);

        return (
          <div key={dir} style={{ marginBottom: 20, padding: 12, background: 'var(--surface-muted)', borderRadius: 8 }}>
            <div style={{ marginBottom: 8, display: 'flex', alignItems: 'center', gap: 8 }}>
              <span style={{ fontWeight: 600 }}>
                {DIR_LABELS[dir] || dir + '方向'}
              </span>
              <span style={{
                fontSize: 11,
                padding: '1px 8px',
                borderRadius: 10,
                background: analysis.monotonic ? '#f6ffed' : '#fff2f0',
                color: analysis.monotonic ? '#52c41a' : '#ff4d4f',
                border: '1px solid ' + (analysis.monotonic ? '#b7eb8f' : '#ffccc7'),
              }}>
                {analysis.monotonic ? 'Rating值连续' + (analysis.trend === 'up' ? '增大' : '减小') : '存在非单调点'}
              </span>
            </div>
            <ResponsiveContainer width="100%" height={200}>
              <LineChart data={points} margin={{ top: 8, right: 24, left: 0, bottom: 8 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="var(--border-color)" />
                <XAxis
                  dataKey="level"
                  tick={{ fontSize: 11 }}
                  angle={-30}
                  textAnchor="end"
                  height={50}
                />
                <YAxis
                  reversed
                  domain={[minVal - padding, maxVal + padding]}
                  tick={{ fontSize: 11 }}
                  label={{ value: 'Rating dB', position: 'insideTopRight', offset: -8, style: { fontSize: 11 } }}
                />
                <Tooltip
                  formatter={function(value, name, props) {
                    var count = props && props.payload && props.payload.count;
                    return [value.toFixed(2) + ' dB' + (count > 1 ? ' (均值, ' + count + '项)' : '')];
                  }}
                  labelFormatter={function(label) { return '等级: ' + label + '（rating 越小越响）'; }}
                />
                <Line
                  type="monotone"
                  dataKey="value"
                  stroke="#1677ff"
                  strokeWidth={2}
                  dot={{ r: 4, fill: '#1677ff' }}
                  activeDot={{ r: 6 }}
                />
                {analysis.violations.map(function(v, i) {
                  return (
                    <ReferenceDot
                      key={i}
                      x={v.level}
                      y={v.value}
                      r={5}
                      fill="#ff4d4f"
                      stroke="#fff"
                      strokeWidth={2}
                    />
                  );
                })}
              </LineChart>
            </ResponsiveContainer>
            <div style={{ marginTop: 6, fontSize: 11, color: 'var(--text-light)' }}>
              共 {points.length} 个等级
              {points.length < 7 ? '（注意：标准测试应覆盖8个等级 MAX ~ MAX-7(MIN)，部分等级数据缺失）' : ''}
            </div>
            {analysis.violations.length > 0 && (
              <div style={{ marginTop: 4, fontSize: 11, color: '#ff4d4f' }}>
                异常点: {analysis.violations.map(function(v) {
                  return v.level + ' (' + (v.diff > 0 ? '+' : '') + v.diff.toFixed(2) + 'dB)';
                }).join(', ')}
              </div>
            )}
            {analysis.monotonic && (
              <div style={{ marginTop: 4, fontSize: 11, color: '#52c41a' }}>
                {analysis.trend === 'up'
                  ? 'Rating 值随等级单调增大，表示实际响度随 MAX → MIN 逐步减小，趋势正常。'
                  : 'Rating 值随等级单调减小，表示实际响度随 MAX → MIN 逐步增大，趋势正常。'}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}
