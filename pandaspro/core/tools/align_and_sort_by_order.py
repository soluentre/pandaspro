def align_and_sort_by_order(
    df,
    input_col,
    order,
    output_col='grade',
    index=True
):
    df = df.copy()

    df[output_col] = (
        df[input_col]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    full_df = pd.DataFrame({output_col: order})

    out = full_df.merge(
        df.drop(columns=[input_col]),
        on=output_col,
        how='left'
    )

    out[output_col] = pd.Categorical(
        out[output_col],
        categories=order,
        ordered=True
    )
    out = out.sort_values(output_col)

    if index:
        out = out.set_index(output_col)

    return out