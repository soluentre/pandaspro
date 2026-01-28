import pandas as pd


def align_and_sort_by_order(
    df,
    input_col,
    order,
    output_col=None,
    index=True
):
    df = df.copy()
    if output_col is None:
        output_col = input_col

    df_key = (
        df[input_col]
        .astype(str)
        .str.strip()
        .str.upper()
    )
    order_key = (
        pd.Series(order)
        .astype(str)
        .str.strip()
        .str.upper()
        .values
    )

    full_df = pd.DataFrame({output_col: order, '__order_key': order_key})
    df = df.drop(columns=[input_col])
    df['__order_key'] = df_key

    out = full_df.merge(
        df,
        on='__order_key',
        how='left'
    )
    out = out.drop(columns=['__order_key'])

    out[output_col] = pd.Categorical(
        out[output_col],
        categories=order,
        ordered=True
    )
    out = out.sort_values(output_col)

    if index:
        out = out.set_index(output_col)

    return out


def ensure_columns(df, cols, fill_value=None):
    df = df.copy()
    for col in cols:
        if col not in df.columns:
            df[col] = fill_value
    return df